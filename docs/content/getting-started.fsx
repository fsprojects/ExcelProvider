(*** hide ***)
#I "../../bin/ExcelProvider"

(**
Getting Started
===========================

To get started simply add ExcelProvider.dll as a reference to your project.

If you are using F# scripts then just add the dll using the `#r` option.

If you then open `FSharp.Interop.Excel` you will have access to the Type Provider functionality.

You can then create a type for an individual workbook. The simplest option is to specify just the name of the workbook. 
You will then be given typed access to the data held in the first sheet. 
The first row of the sheet will be treated as field names and the subsequent rows will be treated as values for these fields.

Parameters
----------

When creating the type you can specify the following parameters:

* `FileName` Location of the Excel file.
* `SheetName` Name of sheet containing data. Defaults to first sheet.
* `Range` Specification using `A1:D3` type addresses of one or more ranges. Defaults to use whole sheet.
* `HasHeaders` Whether the range contains the names of the columns as its first line.
* `ForceString` Specifies forcing data to be processed as strings. Defaults to `false`.

All but the first are optional. 

The parameters can be specified by position or by using the name - for example the following are equivalent:
*)

type MultipleSheets1 = ExcelFile<"MultipleSheets.xlsx", "B">
type MultipleSheets2 = ExcelFile<"MultipleSheets.xlsx", SheetName="B">


(**
Example
-------

This example shows the use of the type provider in an F# script on a sheet containing three rows of data:

![alt text](img/DataTypes.png "Excel sample file with different types")

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type DataTypesTest = ExcelFile<"DataTypes.xlsx">
let file = new DataTypesTest()
let row = file.Data |> Seq.head
let test = row.Float
(** And the variable `test` has the following value: *)
(*** include-value: test ***)