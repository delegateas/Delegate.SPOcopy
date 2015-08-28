#r @"System.ServiceModel"
#r @"Microsoft.IdentityModel"
#r @"Delegate.SPOcopy.dll"

(** Open libraries for use *)
open System
open Delegate.SPOcopy

(**
Unittest setup
==============

Setup of shared values *)
let usr = @"admin@sharepointonlinecopy.onmicrosoft.com"
let pwd =  @"pass@word1"
let km = @"D:\.tmp\spo-copy"
let host = Uri(@"https://sharepointonlinecopy.sharepoint.com")
let url = Uri(host.ToString() + @"Shared%20Documents")
let o365 = 
    Office365.getCookieContainer host usr pwd,
    Office365.userAgent

(**
Test cases
==========

Define *)
let tc1 () = SPOcopy.copy km url usr pwd = ()

(**
Run test cases
==============

Place test cases result in a list *)
let unitTest  = [| tc1; |]
let unitTest' = unitTest |> Array.Parallel.map(fun x -> x())

(** Combine results *)
let result = unitTest' |> Array.reduce ( && )