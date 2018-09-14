(*** hide ***)
#I "../../bin"

(**
ExcelProvider
===========================

This library is for the .NET platform implementing a Excel type provider.

<div class="row">
  <div class="span1"></div>
  <div class="span6">
    <div class="well well-small" id="nuget">
      The library can be <a href="https://nuget.org/packages/ExcelProvider">installed from NuGet</a>:
      <pre>PM> Install-Package ExcelProvider</pre>
    </div>
  </div>
  <div class="span1"></div>
</div>

Example
-------

This example demonstrates the use of the type provider:

![alt text](img/DataTypes.png "Excel sample file with different types")

*)

// reference the type provider dll
#r "ExcelProvider.dll"
open FSharp.Interop.Excel


// Let the type provider do it's work
type DataTypesTest = ExcelFile<"DataTypes.xlsx">
let file = new DataTypesTest()
let row = file.Data |> Seq.head

(**

Now we have strongly typed access to the Excel rows:

![alt text](img/TypedExcel.png "Typed Excel sample file")

*)

row.String
// [fsi:val it : string = "A"]
row.Float
// [fsi:val it : float = 1.0]
row.Boolean
// [fsi:val it : bool = true]


(**

Documentation
-----------------------

For more information see the Documentation pages: 

 * [Getting Started](getting-started.html) contains an overview of the library.
 * [Accessing Sheets](sheets.html) shows how to access different sheets in a workbook.
 * [Accessing Rows](rows.html) shows how to access individual rows in a worksheet.
 * [Accessing Cells](cells.html) shows how to access individual cells within a row of a worksheet.
 * [Accessing Ranges](ranges.html) shows how to access multiple ranges of data within a worksheet.
 * [Without Headers](headers.html) shows how to process sheets which do not include headers.
 

Contributing and copyright
--------------------------

The project is hosted on [GitHub][gh] where you can [report issues][issues], fork 
the project and submit pull requests. If you're adding new public API, please also 
consider adding [samples][content] that can be turned into a documentation. You might
also want to read [library design notes][readme] to understand how it works.

The library is available under Public Domain license, which allows modification and 
redistribution for both commercial and non-commercial purposes. For more information see the 
[License file][license] in the GitHub repository. 

  [content]: https://github.com/fsprojects/ExcelProvider/tree/master/docs/content
  [gh]: https://github.com/fsprojects/ExcelProvider
  [issues]: https://github.com/fsprojects/ExcelProvider/issues
  [readme]: https://github.com/fsprojects/ExcelProvider/blob/master/README.md
  [license]: https://github.com/fsprojects/ExcelProvider/blob/master/LICENSE.txt
*)
