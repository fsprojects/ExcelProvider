module FSharp.Interop.Excel.Tests.ExcelProviderTests

open NUnit.Framework
open FSharp.Interop.Excel
open FsUnit

open System
open System.Collections.Generic
open System.IO

open FSharp.Compiler.Interactive.Shell
open FSharp.Compiler.Tokenization

module FsiTestContext =
    open System.Text

    // Initialize output and input streams
    let sbOut = new StringBuilder()
    let sbErr = new StringBuilder()
    let inStream = new StringReader("")
    let outStream = new StringWriter(sbOut)
    let errStream = new StringWriter(sbErr)

    // Build command line arguments & start FSI session
    let argv = [| "C:\\fsi.exe" |]

    let fi = 
        
#if DEBUG
        FileInfo (__SOURCE_DIRECTORY__ + "/../../src/ExcelProvider.Runtime/bin/Debug/netstandard2.0/ExcelProvider.Runtime.dll")
#else
        FileInfo (__SOURCE_DIRECTORY__ + "/../../src/ExcelProvider.Runtime/bin/Release/netstandard2.0/ExcelProvider.Runtime.dll")

#endif
    let allArgs = Array.append argv [| "--noninteractive"; $"-r:{fi.FullName}" |]
        
    let fsiConfig =
        FsiEvaluationSession.GetDefaultConfiguration()

    let fsiSession =
        FsiEvaluationSession.Create(fsiConfig, allArgs, inStream, outStream, errStream)


    //let a = 123


type BookTest = ExcelFile<"BookTest.xls", "Sheet1", ForceString=true>
type HeaderTest = ExcelFile<"BookTestWithHeader.xls", Range="A2", ForceString=true>
type MultipleRegions = ExcelFile<"MultipleRegions.xlsx", Range="A1:C5,E3:G5", ForceString=true>
type DifferentMainSheet = ExcelFile<"DifferentMainSheet.xlsx">
type DataTypes = ExcelFile<"DataTypes.xlsx">
type DataTypesHeaderTrue = ExcelFile<"DataTypes.xlsx", HasHeaders=true>
type DataTypesHeaderFalse = ExcelFile<"DataTypes.xlsx", HasHeaders=false, ForceString=true>
type DataTypesNoHeader = ExcelFile<"DataTypesNoHeader.xlsx", HasHeaders=false>
type CaseInsensitive = ExcelFile<"DataTypes.XLSX">
type MultiLine = ExcelFile<"MultilineHeader.xlsx">
type MultipleSheetsFirst = ExcelFile<"MultipleSheets.xlsx", "A">
type MultipleSheetsSecond = ExcelFile<"MultipleSheets.xlsx", "B">
type MultipleSheetsSecondRange = ExcelFile<"MultipleSheets.xlsx", "B", "A2">

[<Test>]
let ``Read Text as String``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    row.String |> should equal "A"

[<Test>]
let ``Read General as Boolean if value is TRUE``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    Assert.That(row.Boolean, Is.True)

[<Test>]
let ``Read Number as Float``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    row.Float |> should equal 1.0

[<Test>]
let ``Read Date as DateTime``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    row.Date |> should equal (new DateTime(2014,1,1))

[<Test>]
let ``Read Time as DateTime``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    let expectedTime = new DateTime(1899, 12, 31, 8, 0, 0)
    row.Time |> should equal expectedTime

[<Test>]
let ``Read Currency as Decimal``() =
    let file = DataTypes()
    let row = file.Data |> Seq.head
    row.Currency |> should equal 100.0M

[<Test>]
let ``Empty Text cell should be null``() =
    let file = DataTypes()
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    blankRow.String |> should be Null

[<Test>]
let ``Read blank General as false if column is boolean``() =
    let file = DataTypes()
    let row = file.Data |> Seq.skip 1 |> Seq.head
    Assert.That(row.Boolean, Is.False)

[<Test>]
let ``Empty Number cell should be zero``() =
    let file = DataTypes()
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    let defaultValue = blankRow.Float
    defaultValue |> should equal 0.0f

[<Test>]
let ``Empty Currency cell should be zero``() =
    let file = DataTypes()
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    let defaultValue = blankRow.Currency
    defaultValue |> should equal 0.0M

[<Test>]
let ``Empty date cell should be BOT``() =
    let file = DataTypes()
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    let defaultValue = blankRow.Date
    defaultValue |> should equal DateTime.MinValue

    
[<Test>]
let ``Empty time cell should be today``() =
    let file = DataTypes()
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    let defaultValue = blankRow.Time
    defaultValue |> should equal DateTime.MinValue

[<Test>]
let ``Default Sheet not named Sheet1``() =
    let file = DifferentMainSheet()
    let firstRow = file.Data |> Seq.head
    firstRow.Animal |> should equal "Daisy"
    firstRow.``Pounds of Milk`` |> should equal 12

let expectedToString =
    String.Join (
        Environment.NewLine,
        [ @"Row 1"; "	Animal = Daisy"; "	Pounds of Milk = 12" ])

[<Test>]
let ``ToString format``() =
    let file = DifferentMainSheet()
    let firstRow = file.Data |> Seq.head

    printfn "%O" firstRow

    string firstRow |> should equal expectedToString

[<Test>]
let ``Can get row cell value by column header``() =
    let file = BookTest()
    let row = file.Data |> Seq.head
    row.GetValue "SEC" |> should equal "ASI"

[<Test>]
let ``GetValue with column header is case-sensitive``() =
    let file = BookTest()
    let row = file.Data |> Seq.head
    row.GetValue "SEC" |> should equal "ASI"
    (fun () -> row.GetValue "SeC" |> ignore) |> should throw typeof<Exception>

[<Test>]
let ``GetValue with negative column index should fail``() =
    let file = BookTest()
    let row = file.Data |> Seq.head
    (fun () -> row.GetValue -1 |> ignore) |> should throw typeof<Exception>

[<Test>]
let ``GetValue with column index out of range should be null`` () =
    let file = BookTest()
    let row = file.Data |> Seq.head
    row.GetValue (Int32.MaxValue) |> should equal null

[<Test>]
let ``GetValue with nonexistent column header should fail``() =
    let file = BookTest()
    let row = file.Data |> Seq.head
    (fun () -> row.GetValue "bad header" |> ignore) |> should throw typeof<Exception>

[<Test>]
let ``Can access first row in typed excel data``() =
    let file = BookTest()
    let row = file.Data |> Seq.head
    row.SEC |> should equal "ASI"
    row.BROKER |> should equal "TFS Derivatives HK"

[<Test>]
let ``Can pick an arbitrary header row``() =
    let file = HeaderTest()
    let row = file.Data |> Seq.head
    row.SEC |> should equal "ASI"
    row.BROKER |> should equal "TFS Derivatives HK"

[<Test>]
let ``Can load data from spreadsheet``() =
    let file = Path.Combine(Environment.CurrentDirectory, "BookTestDifferentData.xls")

    printfn "%s" file

    let otherBook = BookTest(file)
    let row = otherBook.Data |> Seq.head

    row.SEC |> should equal "TASI"
    row.STYLE |> should equal "B"
    row.``STRIKE 1`` |> should equal "3"
    row.``STRIKE 2`` |> should equal "4"
    row.``STRIKE 3`` |> should equal "5"
    row.VOL |> should equal "322"

[<Test>]
let ``Can load data from stream``() =
    let file = Path.Combine(Environment.CurrentDirectory, "BookTestDifferentData.xls")

    printfn "%s" file

    use stream = new FileStream(file, FileMode.Open)
    let otherBook = BookTest(stream, ExcelFormat.Binary)
    let row = otherBook.Data |> Seq.head

    row.SEC |> should equal "TASI"
    row.STYLE |> should equal "B"
    row.``STRIKE 1`` |> should equal "3"
    row.``STRIKE 2`` |> should equal "4"
    row.``STRIKE 3`` |> should equal "5"
    row.VOL |> should equal "322"

[<Test>]
let ``Can load from multiple ranges``() =
    let file = MultipleRegions()
    let rows = file.Data |> Seq.toArray

    rows.[0].First |> should equal "A1"
    rows.[0].Second |> should equal "A2"
    rows.[0].Third |> should equal "A3"
    rows.[0].Fourth |> should equal "B1"
    rows.[0].Fifth |> should equal "B2"
    rows.[0].Sixth |> should equal "B3"

    rows.[1].First |> should equal "A4"
    rows.[1].Second |> should equal "A5"
    rows.[1].Third |> should equal "A6"
    rows.[1].Fourth |> should equal "B4"
    rows.[1].Fifth |> should equal "B5"
    rows.[1].Sixth |> should equal "B6"

    rows.[2].First |> should equal "A7"
    rows.[2].Second |> should equal "A8"
    rows.[2].Third |> should equal "A9"
    rows.[2].Fourth |> should equal null
    rows.[2].Fifth |> should equal null
    rows.[2].Sixth |> should equal null

    rows.[3].First |> should equal "A10"
    rows.[3].Second |> should equal "A11"
    rows.[3].Third |> should equal "A12"
    rows.[3].Fourth |> should equal null
    rows.[3].Fifth |> should equal null
    rows.[3].Sixth |> should equal null

[<Test>]
let ``Can load from multiple sheets - first``() =
    let file = MultipleSheetsFirst()
    let rows = file.Data |> Seq.toArray

    rows.[0].First |> should equal 1.0
    rows.[0].Second |> should equal false
    rows.[0].Third |> should equal "a"

    rows.[1].First |> should equal 2.0
    rows.[1].Second |> should equal true
    rows.[1].Third |> should equal "b"

[<Test>]
let ``Can load from multiple sheets - second``() =
    let file = MultipleSheetsSecond()
    let rows = file.Data |> Seq.toArray

    rows.[0].Fourth |> should equal 2.2
    rows.[0].Fifth |> should equal (new DateTime(2013,1,1))

    rows.[1].Fourth |> should equal 3.2
    rows.[1].Fifth |> should equal (new DateTime(2013,2,1))

[<Test>]
let ``Can load data with schema B from specified file and specified sheetname (invalid type data should throw exception)``() =
    let file = MultipleSheetsSecond(__SOURCE_DIRECTORY__ + "/MultipleSheets.xlsx", "A")
    let rows = file.Data |> Seq.toArray

    rows.[0].Fourth |> should equal 1.0
    try
        let _ = rows.[0].Fifth 
        failwith "should not run to here."
    with
    | exn ->
        exn.Message.Split("\n")[0] |> should equal "ExcelProvider: Cannot cast 'false' a 'System.Boolean' to 'System.DateTime'."
    

[<Test>]
let ``Cannot create a type referring to a non-existant sheet at runtime``() =
    // This test is testing runtime behavior, so we need to use the FSI session to run the code
    //printfn "__SOURCE_DIRECTORY__ %s" __SOURCE_DIRECTORY__
    let multiSheet = __SOURCE_DIRECTORY__.Replace(@"\", "/") + "/MultipleSheets.xlsx"
    let result, diag =
        FsiTestContext.fsiSession.EvalInteractionNonThrowing $"""
        open FSharp.Interop.Excel
        type MultipleSheetsThird = ExcelFile<"{multiSheet}", "C">
        //let file = MultipleSheetsThird()
        ()
        """
    //match result with
    //| Choice1Of2 v -> printfn "1 %A" v
    //| Choice2Of2 v -> printfn "2 %A" v
    //printfn "%A" diag
    let dstr = diag[0].ToString()
    dstr.Substring(dstr.IndexOf "typecheck") |> should equal "typecheck error The type provider 'FSharp.Interop.Excel.ExcelProvider.ProviderImplementation+ExcelProvider' reported an error: Sheet [C] does not exist."
    

[<Test>]
let ``Can load from multiple sheets with range``() =
    let file = MultipleSheetsSecondRange()
    let rows = file.Data |> Seq.toArray

    rows.[0].``2.2`` |> should equal 3.2

[<Test>]
let ``Can load file with different casing``() =
    let file = CaseInsensitive()
    // just do one of the same tests as was done for the book we are basing this off of
    let blankRow = file.Data |> Seq.skip 1 |> Seq.head
    let defaultValue = blankRow.Time
    defaultValue |> should equal DateTime.MinValue

[<Test>]
let ``Multiline column name should be converted to a single line``() =
    let file = MultiLine()
    let rows = file.Data |> Seq.toArray
    rows.[0].``Multiline\nheader`` |> should equal "foo"

[<Test>]
let ``HashHeader defaults true``() =
    let file = DataTypesHeaderTrue()
    let row = file.Data |> Seq.head
    row.String |> should equal "A"
    Assert.That(row.Boolean, Is.True)
    row.Float |> should equal 1.0
    row.Date |> should equal (new DateTime(2014,1,1))
    let expectedTime = new DateTime(1899, 12, 31, 8, 0, 0)
    row.Time |> should equal expectedTime
    row.Currency |> should equal 100.0M

[<Test>]
let ``HashHeader false - first row as data``() =
    let file = DataTypesHeaderFalse()
    let row = file.Data |> Seq.head
    let row2 = file.Data |> Seq.skip 1 |> Seq.head
    row.Column1 |> should equal "String"
    row.Column2 |> should equal "Float"
    row2.Column1 |> should equal "A"
    row2.Column2 |> should equal "1"

[<Test>]
let ``HashHeader false with header removed``() =
    let file = DataTypesNoHeader()
    let row = file.Data |> Seq.head
    row.Column1 |> should equal "A"
    row.Column2 |> should equal 1.0
    Assert.That(row.Column3, Is.True)
    row.Column4 |> should equal (new DateTime(2014,1,1))
    let expectedTime = new DateTime(1899, 12, 31, 8, 0, 0)
    row.Column5 |> should equal expectedTime
    row.Column6 |> should equal 100.0M

module AutomaticCoercion = 
  type MixedDataTypes = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes">

  type DateTimeFromTextNumeric = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="H2:H3", HasHeaders=false>
  type DateTimeFromTextString = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="H4:H5", HasHeaders=false>
  type DateTimeFromNumeric = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="H6:H7", HasHeaders=false>
  type DateTimeFromGeneralNumeric = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="H8:H9", HasHeaders=false>

  type DoubleCoercionErrorFromNumberWords = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="B2:B3", HasHeaders=false>
  type DoubleCoercionErrorFromString = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="B4:B5", HasHeaders=false>
  type DoubleCoercionErrorFromCurrencyNumber = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="B6:B7", HasHeaders=false>

  type DateCoercionErrorFromDateWords = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="H2:H3", HasHeaders=false>
  type DateCoercionErrorFromString = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="H4:H5", HasHeaders=false>
  type DateCoercionErrorFromNegativeNumber = ExcelFile<"MixedDataTypes.xlsx", "InvalidDataTypes", Range="H6:H7", HasHeaders=false>


  // See https://github.com/fsprojects/ExcelProvider/issues/14
  type MixedStringTypes = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="A1:A10">
  [<Test>]
  let ``Automatically coerces non-string data in string columns to strings`` () =


      let file = MixedStringTypes()
      let writeTitles data = 
          for (row:MixedStringTypes.Row) in data do 
              (sprintf "%s" row.Title) |> ignore
      (fun () -> writeTitles file.Data) |> should (not' << throw) typeof<System.InvalidCastException>

  type MixedDoubleTypes = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="B1:B10">
  [<Test>]
  let ``Coerces valid string data in numeric columns to double`` () =
      let file = MixedDoubleTypes()
      let getYear data = 
          for (row:MixedDoubleTypes.Row) in data do 
              sprintf "%f" (row.Year) |> ignore
      (fun () -> getYear file.Data) |> should (not' << throw) typeof<System.InvalidCastException>

  type MixedDateTimeTypes = ExcelFile<"MixedDataTypes.xlsx", "ValidDataTypes", Range="H1:H10">
  [<Test>]
  let ``Coerces valid data in datetime columns to double`` () =
      let file = MixedDateTimeTypes()
      let getYear data = 
          for (row:MixedDateTimeTypes.Row) in data do 
              sprintf "%A" (row.``Date Watched``) |> ignore
      (fun () -> getYear file.Data) |> should (not' << throw) typeof<System.InvalidCastException>

  [<Test>]
  let ``Throws when coercing invalid data in to double`` () =
      let filenw = DoubleCoercionErrorFromNumberWords()
      let files = DoubleCoercionErrorFromString()
      let filecn = DoubleCoercionErrorFromCurrencyNumber()
      let getYear data = 
          for (row: DoubleCoercionErrorFromNumberWords.Row) in data do 
              sprintf "%f" (row.``Column1``) |> ignore
      (fun () -> getYear filenw.Data) |> should throw typeof<System.InvalidCastException>

      let getYear data = 
          for (row: DoubleCoercionErrorFromString.Row) in data do 
              sprintf "%f" (row.``Column1``) |> ignore
      (fun () -> getYear files.Data) |> should throw typeof<System.InvalidCastException>

      let getYear data = 
          for (row: DoubleCoercionErrorFromCurrencyNumber.Row) in data do 
              sprintf "%f" (row.``Column1``) |> ignore
      (fun () -> getYear filecn.Data) |> should throw typeof<System.InvalidCastException>


  [<Test>]
  let ``Throws when coercing invalid data to datetime`` () =
      let filedw = DateCoercionErrorFromDateWords()
      let files = DateCoercionErrorFromString()
      let filenn = DateCoercionErrorFromNegativeNumber()
      let getYear data = 
          for (row: DateCoercionErrorFromDateWords.Row) in data do 
              sprintf "%A" (row.``Column1``) |> ignore
      (fun () -> getYear filedw.Data) |> should throw typeof<System.InvalidCastException>

      let getYear data = 
          for (row: DateCoercionErrorFromString.Row) in data do 
              sprintf "%A" (row.``Column1``) |> ignore
      (fun () -> getYear files.Data) |> should throw typeof<System.InvalidCastException>

      let getYear data = 
          for (row: DateCoercionErrorFromNegativeNumber.Row) in data do 
              sprintf "%A" (row.``Column1``) |> ignore
      (fun () -> getYear filenn.Data) |> should throw typeof<System.InvalidCastException>


  [<Test>]
  let ``Automatic conversions do not cause InvalidCastExceptions`` () =
      let file = MixedDataTypes() 
      let printTitles data = 
          for (row:MixedDataTypes.Row) in data do 
              sprintf "%s (%f) %i" row.Title (row.Year) (row.``Date Watched``.Year) |> ignore
      (fun () -> printTitles file.Data) |> should (not' << throw) typeof<System.InvalidCastException>