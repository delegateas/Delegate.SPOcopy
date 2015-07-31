namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("Delegate.SPOcopy")>]
[<assembly: AssemblyProductAttribute("Delegate.SPOcopy")>]
[<assembly: AssemblyDescriptionAttribute("A library that copies a local folder, including files (recursively), to SharePoint Online, ensuring valid SharePoint relative url names.")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
