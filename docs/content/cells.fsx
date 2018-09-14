(*** hide ***)
#I "../../bin/ExcelProvider"

(**
Accessing Cells
===========================

To access a particular cell you need to access the relevant row and use the field name which is the value in the first row of the relevant column.


Example
-------

![alt text](img/MultSheets.png "Excel sample file with multiple sheets")

This example demonstrates referencing the first column (with name `Fourth`) on row 4:

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type MultipleSheetsSecond = ExcelFile<"MultipleSheets.xlsx", "B">
let file = new MultipleSheetsSecond()
let rows = file.Data |> Seq.toArray
let test = rows.[2].Fourth
(** And the variable `test` has the following value: *)
(*** include-value: test ***)

(** Cells can be accessed dynamically using the zero-based column index or case-sensitive header name: *)
let testByIndex = rows.[2].GetValue 0
let testByHeader = rows.[2].GetValue "Fourth"
(** The variables `testByIndex` and `testByHeader` have the respective values: *)
(*** include-value: testByIndex ***)
(*** include-value: testByHeader ***)

(** Accessing cell values by index or string header sacrifices type safety; the result signature is `obj`. *)