(*** hide ***)
// This block of code is omitted in the generated HTML documentation. Use 
// it to define helpers that you do not want to show in the documentation.
#I "../../bin"
#I "../../bin/Delegate.SPOcopy"
#r "Delegate.SPOcopy.dll"

(**
Delegate.SPOcopy
================

<div class="row">
  <div class="span1"></div>
  <div class="span6">
    <div class="well well-small" id="nuget">
      The Delegate.SPOcopy library can be
      <a href="https://nuget.org/packages/Delegate.SPOcopy">installed from NuGet</a>:
      <pre>PM> Install-Package Delegate.SPOcopy</pre>
    </div>
  </div>
  <div class="span1"></div>
</div>

What is it?
-----------

Delegate SharePoint Online copy (Delegate.SPOcopy) is a library that copies a
local folder, including files and subfolders (recursively), to SharePoint Online,
ensuring valid SharePoint relative url names.

Example
-------

Delegate.SPOcopy.Sample.fsx:
*)

open System
open Delegate
open Delegate.SPOcopy

let domain = @"sharepointonlinecopy"
let usr = sprintf @"admin@%s.onmicrosoft.com" domain
let pwd = @"pass@word1"
let source = @"D:\tmp\spo-copy"
let host = Uri( sprintf @"https://%s.sharepoint.com" domain)
let target = Uri(host.ToString() + @"Shared%20Documents")
let o365 = 
    Office365.getCookieContainer host usr pwd,
    Office365.userAgent

copy source target usr pwd LogLevel.Info

(**
Evaluates to the following output when called as:
<pre>
fsianycpu Delegate.SPOcopy.Sample.fsx > Delegate.SPOcopy.Output.txt 2> Delegate.SPOcopy.Error.txt
</pre>

Delegate.SPOcopy.Output.txt:
<pre>
2016-02-03T08:08:01.5752722+01:00 - Info: "SharePoint Online copy (SPOcopy) - Started"
2016-02-03T10:12:27.6846710+01:00 - Info: "SharePoint Online copy (SPOcopy) - Finished"
</pre>

Delegate.SPOcopy.Error.txt:
<pre>
</pre>

> **Remark**: As we can see, the 33.198 files and 1.630 folders (1 GB) are
> created in about two hours on the SharePoint Online instance, with no errors
> whatsoever *)

(** <div align="center"><img src="img/files_local_vs_sp.png" width="75%" height="75%" /></div> *)

(**
How it works and limitations
----------------------------

 * A few words on how the `Delegate.SPOcopy` library works:
    - In order to be able to upload the files with the mminimal amount of 
      **noise** we rely on using SharePoint Onlines REST service instead of their
      SOAP, see [Martin Lawrence REST vs SOAP][ml] for more info. This allows
      us to use F# powerfull `async` but simple engine to implement parallelism.
    - For more information, please look into the code (about +275 lines) at [GitHub][gh]

 * We describe a few **limitations** we found while we were making the library:
    - **Only works with Office365 account without ADFS**: As we have borrowed
      [Ronnie Holms][rh] Office365 module, we haven't expanded it so it also
      would supports ADFS users as we think the easiest approach is that the
      Office365 Admin just creates a new service account in the cloud without
      synchronization with the local AD.
    - **No executable**: The reason we haven't created an executable file is that
      we then have to rely on .bat or .cmd files in order to execute the 
      application command-line arguments. We think that the approach of creating
      a type-safe F# script file is a much better approach.
*)
 
(**
Contributing and copyleft
--------------------------

The project is hosted on [GitHub][gh] where you can [report issues][issues], fork 
the project and submit pull requests.

The library is available under an Open Source MIT license, which allows modification and 
redistribution for both commercial and non-commercial purposes. For more information see the 
[License file][license] in the GitHub repository. 

  [ml]: http://stackoverflow.com/a/8983122
  [rh]: https://twitter.com/ronnieholm
  [gh]: https://github.com/delegateas/Delegate.SPOcopy
  [issues]: https://github.com/delegateas/Delegate.SPOcopy/issues
  [license]: https://github.com/delegateas/Delegate.SPOcopy/blob/sync/LICENSE.md 
*)
