module ExcelProvider.ExcelProvider

open System
open System.Collections.Generic
open System.IO
open System.Reflection

open ExcelProvider.Helper
open ExcelProvider.ExcelAddressing
open ICSharpCode.SharpZipLib.Zip
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes

// Represents a row in a provided ExcelFileInternal
type Row(rowIndex, getCellValue: int -> int -> obj, columns: Map<string, int>) =
    member this.GetValue columnIndex = getCellValue rowIndex columnIndex

    override this.ToString() =
        let columnValueList =
            [for column in columns do
                let value = getCellValue rowIndex column.Value
                let columnName, value = column.Key, string value
                yield sprintf "\t%s = %s" columnName value]
            |> String.concat Environment.NewLine

        sprintf "Row %d%s%s" rowIndex Environment.NewLine columnValueList

// Avoids "warning FS0025: Incomplete pattern matches on this expression"
// when using: (fun [row] -> <@@ ... @@>)
let private singleItemOrFail func items = 
    match items with
    | [ item ] -> func item
    | _ -> failwith "Expected single item list."

// Avoids "warning FS0025: Incomplete pattern matches on this expression"
// when using: (fun [] -> <@@ ... @@>)
let private emptyListOrFail func items = 
    match items with
    | [] -> func()
    | _ -> failwith "Expected empty list"


// get the type, and implementation of a getter property based on a template value
let internal propertyImplementation columnIndex (value : obj) =
    match value with
    | :? float -> typeof<double>, (fun row -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> (v :?> Nullable<double>).GetValueOrDefault()) @@>) |> singleItemOrFail
    | :? bool -> typeof<bool>, (fun row -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> (v :?> Nullable<bool>).GetValueOrDefault()) @@>) |> singleItemOrFail
    | :? DateTime -> typeof<DateTime>, (fun row -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> (v :?> Nullable<DateTime>).GetValueOrDefault()) @@>) |> singleItemOrFail
    | :? string -> typeof<string>, (fun row -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> v :?> string) @@>) |> singleItemOrFail
    | _ -> typeof<obj>, (fun row -> <@@ (%%row: Row).GetValue columnIndex @@>) |> singleItemOrFail

// gets a list of column definition information for the columns in a view
let internal getColumnDefinitions (data : View) forcestring =
    let getCell = getCellValue data
    [for columnIndex in 0 .. data.ColumnMappings.Count - 1 do
        let rawColumnName = getCell 0 columnIndex |> string
        if not (String.IsNullOrWhiteSpace(rawColumnName)) then
            let processedColumnName = rawColumnName.Replace("\n", "\\n")
            let cellType, getter =
                if forcestring then
                    let getter = (fun row ->
                                    <@@
                                        let value = (%%row: Row).GetValue columnIndex |> string
                                        if String.IsNullOrEmpty value then null
                                        else value
                                    @@>) |> singleItemOrFail
                    typedefof<string>, getter
                else
                    let cellValue = getCell 1 columnIndex
                    propertyImplementation columnIndex cellValue
            yield (processedColumnName, (columnIndex, cellType, getter))]

// Simple type wrapping Excel data
type ExcelFileInternal(filename, sheetname, range) =

    let data =
        let view = openWorkbookView filename sheetname range
        let columns = [for (columnName, (columnIndex, _, _)) in getColumnDefinitions view true -> columnName, columnIndex] |> Map.ofList
        let buildRow rowIndex = new Row(rowIndex, getCellValue view, columns)
        seq{ 1 .. view.RowCount}
        |> Seq.map buildRow

    member __.Data = data

type internal GlobalSingleton private () =
    static let mutable instance = Dictionary<_, _>()
    static member Instance = instance

let internal memoize f x =
    if (GlobalSingleton.Instance).ContainsKey(x) then (GlobalSingleton.Instance).[x]
    else
        let res = f x
        (GlobalSingleton.Instance).[x] <- res
        res
do
    let loadedAssemblies = new HashSet<string>()
    let sharpZipLibAssemblyName =
        let zipFileType = typedefof<ZipFile>
        zipFileType.Assembly.GetName()
    let resolveAssembly sender (resolveEventArgs : ResolveEventArgs) =
        let assemblyName = resolveEventArgs.Name
        if loadedAssemblies.Add( assemblyName ) then
            if assemblyName.StartsWith(sharpZipLibAssemblyName.Name)
            then Assembly.Load( sharpZipLibAssemblyName.FullName )
            else null
        else null
    let handler = new ResolveEventHandler( resolveAssembly )
    AppDomain.CurrentDomain.add_AssemblyResolve handler

let internal typExcel(cfg:TypeProviderConfig) =
    let executingAssembly = System.Reflection.Assembly.GetExecutingAssembly()

    // Create the main provided type
    let excelFileProvidedType = ProvidedTypeDefinition(executingAssembly, rootNamespace, "ExcelFile", Some(typeof<ExcelFileInternal>))

    /// Given a function to format names (such as `niceCamelName` or `nicePascalName`)
    /// returns a name generator that never returns duplicate name (by appending an
    /// index to already used names)
    /// 
    /// This function is curried and should be used with partial function application:
    ///
    ///     let makeUnique = uniqueGenerator nicePascalName
    ///     let n1 = makeUnique "sample-name"
    ///     let n2 = makeUnique "sample-name"
    ///
    let uniqueGenerator niceName =
      let set = new HashSet<_>()
      fun name ->
        let mutable name = niceName name
        while set.Contains name do 
          let mutable lastLetterPos = String.length name - 1
          while Char.IsDigit name.[lastLetterPos] && lastLetterPos > 0 do
            lastLetterPos <- lastLetterPos - 1
          if lastLetterPos = name.Length - 1 then
            if name.Contains " " then
                name <- name + " 2"
            else
                name <- name + "2"
          elif lastLetterPos = 0 && name.Length = 1 then
            name <- (UInt64.Parse name + 1UL).ToString()
          else
            let number = name.Substring(lastLetterPos + 1)
            name <- name.Substring(0, lastLetterPos + 1) + (UInt64.Parse number + 1UL).ToString()
        set.Add name |> ignore
        name

    let buildTypes (typeName:string) (args:obj[]) =
        let filename = args.[0] :?> string
        let sheetname = args.[1] :?> string
        let range = args.[2] :?> string
        let forcestring = args.[3] :?> bool

        // resolve the filename relative to the resolution folder
        let resolvedFilename = Path.Combine(cfg.ResolutionFolder, filename)

        let ProvidedTypeDefinitionExcelCall (filename, sheetname, range, forcestring)  =
            let gen = uniqueGenerator id
            let data = openWorkbookView resolvedFilename sheetname range

            // define a provided type for each row, erasing to a int -> obj
            let providedRowType = ProvidedTypeDefinition("Row", Some(typeof<Row>))

            // add one property per Excel field
            let columnProperties = getColumnDefinitions data forcestring
            for (columnName, (columnIndex, propertyType, getter)) in columnProperties do

                let prop = ProvidedProperty(columnName |> gen, propertyType, GetterCode = getter)
                // Add metadata defining the property's location in the referenced file
                prop.AddDefinitionLocation(1, columnIndex, filename)
                providedRowType.AddMember(prop)

            // define the provided type, erasing to an seq<int -> obj>
            let providedExcelFileType = ProvidedTypeDefinition(executingAssembly, rootNamespace, typeName, Some(typeof<ExcelFileInternal>))

            // add a parameterless constructor which loads the file that was used to define the schema
            providedExcelFileType.AddMember(ProvidedConstructor([], InvokeCode = emptyListOrFail (fun () -> <@@ ExcelFileInternal(resolvedFilename, sheetname, range) @@>)))

            // add a constructor taking the filename to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>)], InvokeCode = singleItemOrFail (fun filename -> <@@ ExcelFileInternal(%%filename, sheetname, range) @@>)))

            // add a new, more strongly typed Data property (which uses the existing property at runtime)
            providedExcelFileType.AddMember(ProvidedProperty("Data", typedefof<seq<_>>.MakeGenericType(providedRowType), GetterCode = singleItemOrFail (fun excFile -> <@@ (%%excFile:ExcelFileInternal).Data @@>)))

            // add the row type as a nested type
            providedExcelFileType.AddMember(providedRowType)

            providedExcelFileType

        (memoize ProvidedTypeDefinitionExcelCall)(filename, sheetname, range, forcestring)
    
    let parameters = 
        [ ProvidedStaticParameter("FileName", typeof<string>) 
          ProvidedStaticParameter("SheetName", typeof<string>, parameterDefaultValue = "") 
          ProvidedStaticParameter("Range", typeof<string>, parameterDefaultValue = "") 
          ProvidedStaticParameter("ForceString", typeof<bool>, parameterDefaultValue = false) ]

    let helpText = 
        """<summary>Typed representation of data in an Excel file.</summary>
           <param name='FileName'>Location of the Excel file.</param>
           <param name='SheetName'>Name of sheet containing data. Defaults to first sheet.</param>
           <param name='Range'>Specification using `A1:D3` type addresses of one or more ranges. Defaults to use whole sheet.</param>
           <param name='ForceString'>Specifies forcing data to be processed as strings. Defaults to `false`.</param>"""

    do excelFileProvidedType.AddXmlDoc helpText
    do excelFileProvidedType.DefineStaticParameters(parameters, buildTypes)

    // add the type to the namespace
    excelFileProvidedType

[<TypeProvider>]
type public ExcelProvider(cfg:TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces()

    do this.AddNamespace(rootNamespace,[typExcel cfg])

[<TypeProviderAssembly>]
do ()
