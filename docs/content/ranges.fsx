(*** hide ***)
#I "../../bin/ExcelProvider"

(**
Accessing Ranges
===========================

To access a range you need to specify the range (or ranges) to be used.

You can either specify this as the second parameter or the named `Range` parameter.


Example
-------

![alt text](img/Excel.png "Excel sample file with multiple ranges")

This example demonstrates referencing multiple ranges:

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel

// Let the type provider do it's work
type MultipleRegions = ExcelFile<"MultipleRegions.xlsx", Range="A1:C5,E3:G5", ForceString=true>
let file = new MultipleRegions()
let rows = file.Data |> Seq.toArray

let test1 = rows.[0].First
let test2 = rows.[0].Fourth
(** And the variables `test1` and `test2` have the following values: *)
(*** include-value: test1 ***)
(*** include-value: test2 ***)