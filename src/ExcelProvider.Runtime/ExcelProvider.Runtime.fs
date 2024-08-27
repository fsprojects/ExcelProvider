namespace FSharp.Interop.Excel

type ExcelFormat =
    | Xlsx
    | Csv
    | Binary

namespace FSharp.Interop.Excel.ExcelProvider

open System
open System.IO
open System.Data
open System.Text.RegularExpressions
open ExcelDataReader
open FSharp.Core.CompilerServices
open FSharp.Interop.Excel

[<AutoOpen>]
module internal ExcelAddressing =

    type Address =
        { Sheet: string; Row: int; Column: int }

    type Range =
        | Bounded of Address * Address
        | Unbounded of Address

    type RangeView =
        { StartColumn: int
          EndColumn: int
          StartRow: int
          EndRow: int
          Sheet: DataTable }

    type View =
        { RowCount: int
          ColumnMappings: Map<int, int * RangeView> }

    ///Parses an excel row column address of the form <COLUMN LETTERS><ROW NUMBER> into a zero based row, column address.
    ///For example the address A1 would be parsed as 0, 0
    let parseCellAddress cellAddress =

        if String.IsNullOrWhiteSpace cellAddress then
            0, 0
        else
            let convertToBase radix digits =
                let digitValue i digit =
                    float digit * Math.Pow(float radix, float i)

                digits |> List.rev |> List.mapi digitValue |> Seq.sum |> int

            let charToDigit char = ((int) (Char.ToUpper(char))) - 64

            let column =
                cellAddress
                |> Seq.filter Char.IsLetter
                |> Seq.map charToDigit
                |> Seq.toList
                |> convertToBase 26

            let row =
                cellAddress
                |> Seq.filter Char.IsNumber
                |> Seq.map (string >> Int32.Parse)
                |> Seq.toList
                |> convertToBase 10

            (row - 1), (column - 1)

    let addressParser = new Regex("((?<sheet>[^!]*)!)?(?<cell>\w+\d+)?")

    ///Parses an excel address from a string
    ///Valid inputs look like:
    ///Sheet!A1
    ///B3
    let parseExcelAddress sheetContext address =
        if address <> sheetContext then
            let regexMatch = addressParser.Match(address)
            let sheetGroup = regexMatch.Groups.Item("sheet")
            let cellGroup = regexMatch.Groups.Item("cell")

            let sheet =
                if sheetGroup.Success then
                    sheetGroup.Value
                else
                    sheetContext

            let row, column = parseCellAddress cellGroup.Value

            { Sheet = sheet
              Row = row
              Column = column }
        else
            { Sheet = sheetContext
              Row = 0
              Column = 0 }

    ///Parses an excel range of the form
    ///<ADDRESS>:<ADDRESS> | <ADDRESS>
    ///ADDRESS is parsed with the parseExcelAddress function
    let parseExcelRange sheetContext (range: string) =
        let addresses = range.Split(':') |> Array.map (parseExcelAddress sheetContext)

        match addresses with
        | [| a; b |] -> Bounded(a, b)
        | [| a |] -> Unbounded a
        | _ -> failwith (sprintf "ExcelProvider: A range can contain only one or two address [%s]" range)

    ///Parses a potential sequence of excel ranges, seperated by commas
    let parseExcelRanges sheetContext (range: string) =
        range.Split(',') |> Array.map (parseExcelRange sheetContext) |> Array.toList

    ///Gets the start and end offsets of a range
    let getRangeView (workbook: DataSet) range =
        let topLeft, bottomRight, sheet =
            match range with
            | Bounded(topLeft, bottomRight) ->
                let sheet = workbook.Tables.[topLeft.Sheet]
                topLeft, bottomRight, sheet
            | Unbounded(topLeft) ->
                let sheet = workbook.Tables.[topLeft.Sheet]

                topLeft,
                { topLeft with
                    Row = sheet.Rows.Count
                    Column = sheet.Columns.Count - 1 },
                sheet

        { StartColumn = topLeft.Column
          StartRow = topLeft.Row
          EndColumn = bottomRight.Column
          EndRow = bottomRight.Row
          Sheet = sheet }

    ///Gets a View object which can be used to read data from the given range in the DataSet
    let public getView (workbook: DataSet) sheetname range =
        let worksheets = workbook.Tables

        let workSheetName =
            if worksheets.Contains sheetname then
                sheetname
            else if sheetname = null || sheetname = "" then
                worksheets.[0].TableName //accept TypeProvider without specific SheetName...
            else
                failwithf "ExcelProvider: Sheet [%s] does not exist." sheetname

        let ranges =
            parseExcelRanges workSheetName range |> List.map (getRangeView workbook)

        let minRow = ranges |> Seq.map (fun range -> range.StartRow) |> Seq.min
        let maxRow = ranges |> Seq.map (fun range -> range.EndRow) |> Seq.max
        let rowCount = maxRow - minRow

        let rangeViewOffsetRecord rangeView =
            seq { rangeView.StartColumn .. rangeView.EndColumn }
            |> Seq.map (fun i -> i, rangeView)
            |> Seq.toList

        let rangeViewsByColumn =
            ranges |> Seq.map rangeViewOffsetRecord |> Seq.concat |> Seq.toList

        if rangeViewsByColumn |> Seq.distinctBy fst |> Seq.length < rangeViewsByColumn.Length then
            failwith "ExcelProvider: Ranges cannot overlap"

        let columns =
            rangeViewsByColumn |> Seq.mapi (fun index entry -> (index, entry)) |> Map.ofSeq

        { RowCount = rowCount
          ColumnMappings = columns }

    ///Reads the value of a cell from a view
    let getCellValue view row column =
        if column < 0 then
            failwith "ExcelProvider: Column index must be nonnegative"

        let columns = view.ColumnMappings

        match columns.TryFind column with
        | Some(sheetColumn, rangeView) ->
            let row = rangeView.StartRow + row
            let sheet = rangeView.Sheet

            if row < sheet.Rows.Count && sheetColumn < sheet.Columns.Count then
                match rangeView.Sheet.Rows.[row].Item(sheetColumn) with
                | :? System.DBNull -> null
                | nonNullValue -> nonNullValue
            else
                null
        | _ -> null

    ///Reads the contents of an excel file into a DataSet
    let public openWorkbookView filename sheetname range =

#if NETSTANDARD || NETCOREAPP
        // Register encodings
        do System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
#endif
        let fail action (ex: exn) =
            let exceptionTypeName = ex.GetType().Name

            let message =
                sprintf "ExcelProvider: Could not %s. %s - %s" action exceptionTypeName (ex.Message)

            failwith message

        use stream =
            try
                new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            with ex ->
                let action = sprintf "open file '%s'" filename
                fail action ex

        let excelReader =
            let action = "create ExcelDataReader"

            try
                let reader =
                    if filename.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) then
                        ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream)
                    elif filename.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) then
                        ExcelDataReader.ExcelReaderFactory.CreateCsvReader(stream)
                    else
                        ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream)

                if reader.IsClosed then
                    fail
                        action
                        (Exception
                            "ExcelProvider: The reader was closed on startup without raising a specific exception")

                reader
            with ex ->
                fail action ex

        let workbook =
            excelReader.AsDataSet(
                new ExcelDataSetConfiguration(
                    ConfigureDataTable = (fun _ -> new ExcelDataTableConfiguration(UseHeaderRow = false))
                )
            )

        let range =
            if String.IsNullOrWhiteSpace range then
                sheetname //workbook.Tables.[0].TableName <== maybe the root cause of https://github.com/fsprojects/ExcelProvider/issues/77
            else
                range

        let view = getView workbook sheetname range
        (excelReader :> IDisposable).Dispose()
        view

    ///Reads the contents of an excel file into a DataSet
    let public openWorkbookViewFromStream (stream: Stream, format: ExcelFormat) sheetname range =

#if NETSTANDARD || NETCOREAPP
        // Register encodings
        do System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)
#endif
        let fail action (ex: exn) =
            let exceptionTypeName = ex.GetType().Name

            let message =
                sprintf "ExcelProvider: Could not %s. %s - %s" action exceptionTypeName (ex.Message)

            failwith message

        let excelReader =
            let action = "create ExcelDataReader"

            try
                let reader =
                    match format with
                    | Xlsx -> ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream)
                    | Csv -> ExcelDataReader.ExcelReaderFactory.CreateCsvReader(stream)
                    | Binary -> ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream)

                if reader.IsClosed then
                    fail action (Exception "The reader was closed on startup without raising a specific exception")

                reader
            with ex ->
                fail action ex

        let workbook =
            excelReader.AsDataSet(
                new ExcelDataSetConfiguration(
                    ConfigureDataTable = (fun _ -> new ExcelDataTableConfiguration(UseHeaderRow = false))
                )
            )

        let range =
            if String.IsNullOrWhiteSpace range then
                sheetname //workbook.Tables.[0].TableName <== maybe the root cause of https://github.com/fsprojects/ExcelProvider/issues/77
            else
                range

        let view = getView workbook sheetname range
        (excelReader :> IDisposable).Dispose()
        view

    let failInvalidCast fromObj (fromType: Type) (toType: Type) columnName rowIndex filename sheetname =
        sprintf
            "ExcelProvider: Cannot cast '%A' a '%s' to '%s'.\nFile: '%s'. Sheet: '%s'\nColumn '%s'. Row %i."
            fromObj
            fromType.FullName
            toType.FullName
            filename
            sheetname
            columnName
            rowIndex
        |> InvalidCastException
        |> raise

    // gets a list of column definition information for the columns in a view
    let internal getColumnDefinitions (data: View) hasheaders =
        let getCell = getCellValue data

        [ for columnIndex in 0 .. data.ColumnMappings.Count - 1 do
              let rawColumnName =
                  if hasheaders then
                      getCell 0 columnIndex |> string
                  else
                      "Column" + (columnIndex + 1).ToString()

              if not (String.IsNullOrWhiteSpace(rawColumnName)) then
                  let processedColumnName = rawColumnName.Replace("\n", "\\n")
                  yield (processedColumnName, columnIndex) ]

// Represents a row in a provided ExcelFileInternal
type Row(documentId, sheetname, rowIndex, getCellValue: int -> int -> obj, columns: Map<string, int>) =
    member this.GetValue columnIndex = getCellValue rowIndex columnIndex

    member this.GetValue columnName =
        match columns.TryFind columnName with
        | Some(columnIndex) -> this.GetValue columnIndex
        | None ->
            columns
            |> Seq.map (fun kvp -> kvp.Key)
            |> Seq.tryFind (fun header -> String.Equals(header, columnName, StringComparison.OrdinalIgnoreCase))
            |> function
                | Some header ->
                    sprintf "ExcelProvider: Column \"%s\" was not found. Did you mean \"%s\"?" columnName header
                | None -> sprintf "ExcelProvider: Column \"%s\" was not found." columnName
            |> failwith

    member this.TryGetValue<'a> (columnIndex: int) columnName =
        let value = this.GetValue columnIndex

        try
            value :?> 'a
        with :? InvalidCastException ->
            failInvalidCast value (value.GetType()) typeof<'a> columnName rowIndex documentId sheetname

    member this.TryGetNullableValue<'a when 'a: (new: unit -> 'a) and 'a: struct and 'a :> ValueType>
        (columnIndex: int)
        columnName
        =
        let value = this.GetValue columnIndex

        try
            (value :?> Nullable<'a>).GetValueOrDefault()
        with :? InvalidCastException ->
            failInvalidCast value (value.GetType()) typeof<'a> columnName rowIndex documentId sheetname

    override this.ToString() =
        let columnValueList =
            [ for column in columns do
                  let value = getCellValue rowIndex column.Value
                  let columnName, value = column.Key, string value
                  yield sprintf "\t%s = %s" columnName value ]
            |> String.concat Environment.NewLine

        sprintf "Row %d%s%s" rowIndex Environment.NewLine columnValueList

// Simple type wrapping Excel data
type ExcelFileInternal private (view, documentId, sheetname, hasheaders) =

    let data =
        let columns =
            [ for (columnName, columnIndex) in getColumnDefinitions view hasheaders -> columnName, columnIndex ]
            |> Map.ofList

        let buildRow rowIndex =
            new Row(documentId, sheetname, rowIndex, getCellValue view, columns)

        seq { (if hasheaders then 1 else 0) .. view.RowCount } |> Seq.map buildRow

    new(filename, sheetname, range, hasheaders) =
        let view = openWorkbookView filename sheetname range
        ExcelFileInternal(view, filename, sheetname, hasheaders)

    new(stream, format, sheetname, range, hasheaders) =
        let view = openWorkbookViewFromStream (stream, format) sheetname range
        ExcelFileInternal(view, "stream", sheetname, hasheaders)

    member __.Data = data


module Attributes =

    [<TypeProviderAssembly("ExcelProvider.DesignTime.dll")>]
    do ()
