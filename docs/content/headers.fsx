(*** hide ***)
#I "../../bin/ExcelProvider"

(**
Without Headers
===========================

To process a sheet which does not include headers you can use the `HasHeaders` parameter.

This parameter defaults to `true`.

If you set it to `false`, then all rows are treated as data.


Example
-------

This example shows the use of the type provider in an F# script on a sheet containing no headers:

![alt text](img/DataTypesNoHeader.png "Excel sample file with different types and no headers")

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type DataTypesTest = ExcelFile<"DataTypesNoHeader.xlsx", HasHeaders=false>
let file = new DataTypesTest()
let row = file.Data |> Seq.head
let test = row.Column2
(** And the variable `test` has the following value: *)
(*** include-value: test ***)