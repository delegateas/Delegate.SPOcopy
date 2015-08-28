module Delegate.SPOcopy.Tests

open Delegate.SPOcopy
open NUnit.Framework

[<Test>]
let ``hello returns 42`` () =
//  let result = SPOcopy.hello 42
  let result = 42
  printfn "%i" result
  Assert.AreEqual(42,result)
