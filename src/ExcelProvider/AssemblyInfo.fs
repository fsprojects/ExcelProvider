namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("ExcelProvider")>]
[<assembly: AssemblyProductAttribute("ExcelProvider")>]
[<assembly: AssemblyDescriptionAttribute("A type provider for regular expressions.")>]
[<assembly: AssemblyVersionAttribute("0.0.1")>]
[<assembly: AssemblyFileVersionAttribute("0.0.1")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "0.0.1"
