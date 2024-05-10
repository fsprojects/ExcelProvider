(**
---
title: Accessing Sheets
category: Documentation
categoryindex: 2
index: 2
---
*)


(*** hide ***)
#r "../bin/ExcelProvider.Runtime/netstandard2.0/ExcelProvider.Runtime.dll"


(**
Accessing Sheets
===========================

To access a particular sheet you need to specify the sheet name as the second parameter when creating the ExcelProvider type.

If you do not include a second parameter then the first sheet in the workbook is used.

Example
-------

![alt text](images/MultSheets.png "Excel sample file with multiple sheets")

This example demonstrates referencing the second sheet (with name `B`):

*)

// reference the type provider dll
#r "ExcelProvider.Runtime.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type MultipleSheetsSecond = ExcelFile<"MultipleSheets.xlsx", "B">
let file = new MultipleSheetsSecond()
let rows = file.Data |> Seq.toArray
let test = rows.[0].Fourth
(** And the variable `test` has the following value: *)
(*** include-value: test ***)
