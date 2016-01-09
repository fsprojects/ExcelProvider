namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("ExcelProvider")>]
[<assembly: AssemblyProductAttribute("ExcelProvider")>]
[<assembly: AssemblyDescriptionAttribute("This library is for the .NET platform implementing a Excel type provider.")>]
[<assembly: AssemblyVersionAttribute("0.2.2")>]
[<assembly: AssemblyFileVersionAttribute("0.2.2")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "0.2.2"
