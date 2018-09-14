(*** hide ***)
#I "../../bin/ExcelProvider"

(**
Accessing Rows
===========================

Rows are returned as a sequence from the `Data` element of the ExcelFile type.

Example
-------

![alt text](img/MultSheets.png "Excel sample file with multiple sheets")

This example demonstrates loading the second row (with index 1) into the variable test:

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type MultipleSheetsSecond = ExcelFile<"MultipleSheets.xlsx", "B">
let file = new MultipleSheetsSecond()
let rows = file.Data |> Seq.toArray
let test = rows.[1]
(** And the variable `test` has the following value: *)
(*** include-value: test ***)