namespace Delegate.SPOcopy

module SPOcopy = 
  open System
  open System.IO
  open System.Net
  open System.Text.RegularExpressions
  open System.Threading
  open System.Xml.Linq

  open Microsoft.FSharp.Reflection

  type SPOitems = Folders | Files

  [<AutoOpen>]
  module Either = 
    type ('a,'b) Choice with 
      static member Left  x = Choice1Of2 x
      static member Right x = Choice2Of2 x
    type ('a,'b) Either = Choice<'a,'b>

    let (|Left|Right|) = function
      | Choice1Of2 x -> Left x | Choice2Of2 x -> Right x
    let bind f = function | Left a -> f a | Right b -> Either.Right b
    let chooseLeft  = function | Left s -> Some s | Right _ -> None
    let chooseRight = function | Left _ -> None   | Right b -> Some b

  [<AutoOpen>]
  module Logger =
    // Info    = (1 <<< 0) ||| (1 <<< 2)
    // Warning = (1 <<< 1) ||| (1 <<< 2)
    // Error   = (1 <<< 2)
    // Verbose = (1 <<< 0) ||| (1 <<< 1) ||| (1 <<< 2) ||| (1 <<< 3)
    type LogLevel = 
        | Info    =  5 // 1 (2^0) Indicates logs for an informational message.
        | Warning =  6 // 2 (2^1) Indicates logs for a warning.
        | Error   =  4 // 4 (2^2) Indicates logs for an error.
        | Verbose = 15 // 8 (2^3) Indicates logs at all levels.

    let ts () = DateTime.Now.ToString("o") // ISO-8601
    let cw (s:string) = Console.WriteLine(s)
    let cew (s:string) = Console.Error.WriteLine(s)

    let log l x =
        let msg = sprintf "%s - %A: %A" (ts()) l x
        match l with
        | LogLevel.Error | LogLevel.Warning -> cew msg
        | LogLevel.Info  | LogLevel.Verbose -> cw msg
        | _ -> ()

  [<AutoOpen>]
  module Ext = 
    module Seq = 
      let split size source =  seq { 
        let r = ResizeArray()
        for x in source do
          r.Add(x)
          if r.Count = size then 
            yield r.ToArray()
            r.Clear()
        if r.Count > 0 then yield r.ToArray() }
  
  [<AutoOpen>]
  module internal SPO = 

    module Util =
      let ducToString (x:'a) = 
        match FSharpValue.GetUnionFields(x, typeof<'a>) with
        | case, _ -> case.Name

      // Valid SharePoint relative url names
      // - cannot ends with the following strings: .aspx, .ashx, .asmx, .svc,
      // - cannot begin or end with a dot,
      // - cannot contain consecutive dots and
      // - cannot contain any of the following characters: ~ " # % & * : < > ? / \ { | }.
      let escape (name : string) = 
        name
        |> fun x -> 
          Seq.fold (fun (a : string) y -> a.EndsWith(y) |> function
            | false -> a
            | true -> a.Replace(y, "-")) x [ ".aspx";".ashx";".asmx";".svc" ]
        |> fun x -> 
          Seq.fold (fun (a : string) y -> a.Replace(y, "-")) x 
            [ ".aspx";".ashx";".asmx";".svc" ]
        |> fun x -> x.StartsWith(".") |> function
          | false -> x
          | true -> "-" + x.[1..(x.Length - 1)]
        |> fun x -> x.EndsWith(".") |> function
          | false -> x
          | true -> x.[0..(x.Length - 2)] + "-"
        |> fun x -> Regex.Replace(x, "\.{2,}", "-")
        |> fun x -> 
          Seq.fold (fun (a : string) y -> a.Replace(y, '-')) x 
            [ '\\';'/';':';'*';'?';'"';'<';'>';'|';'#';'{';'}';'%';'~';'&' ]

      let nameUrl urlRoot root (path:string) =
        let name = Path.GetFileName path 
        path.Replace(root, String.Empty).Replace(name, String.Empty).Split('\\')
        |> Array.map(escape)
        |> Array.fold(fun a x -> a + @"/" + x) urlRoot
        |> fun releativeUrl -> releativeUrl, name

      let spoUrls host releativeUrl name items = 
        let check =
          sprintf
            "%s_api/web/GetFolderByServerRelativeUrl('%s')/%s?$filter=name eq '%s'"
              host releativeUrl (ducToString items) name
        let create = items |> function
          | Folders -> 
            sprintf
              "%s_api/web/%s/add('%s/%s')" host (ducToString items) releativeUrl name
          | Files ->
            sprintf 
              "%s_api/web/GetFolderByServerRelativeUrl('%s')/%s/add(url='%s',overwrite=true)"
                host releativeUrl (ducToString items) name
        check, create

    module CRUD =
      let formDigest o365 (host:Uri) = 
        let ns = @"http://schemas.microsoft.com/ado/2007/08/dataservices"
        try 
          let cookie, agent = o365
          let url = Uri(host.ToString() + "_api/contextinfo")
          let req = HttpWebRequest.Create(url) :?> HttpWebRequest
          req.Method <- "POST"
          req.ContentLength <- 0L
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          use wresp = req.GetResponse() :?> HttpWebResponse
          use sr = new StreamReader(wresp.GetResponseStream())
          let xml = sr.ReadToEnd()
          let xml' = XDocument.Parse(xml)
          xml'.Descendants()
          |> Seq.where (fun x -> x.Name = XNamespace.Get(ns) + "FormDigestValue")
          |> Seq.head
          |> fun x -> x.Value
        with ex -> log LogLevel.Error ex; raise ex

      let exists (path:string) spoUrls o365 = 
        let url, _ = spoUrls
        let cookie, agent = o365
        async{
          let req = HttpWebRequest.Create(requestUriString = url) :?> HttpWebRequest
          req.Method <- "GET"
          req.Accept <- "application/json;odata=verbose"
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          req.Timeout <- Timeout.Infinite
          use! rsp = req.AsyncGetResponse()
          use stream = rsp.GetResponseStream()
          use reader = new StreamReader(stream)
          let json = reader.ReadToEnd()
          return 
            match json.Contains("\"Exists\": true") with
            | true -> None | false -> Some (path, spoUrls) } |> Async.Catch

      let create (path:string) spoUrls o365 digest items = 
        let _, url = spoUrls
        let cookie, agent = o365
        async{
          let req = HttpWebRequest.Create(requestUriString = url) :?> HttpWebRequest
          req.Headers.Add("X-RequestDigest", digest)
          req.Method <- "POST"
          req.Accept <- "application/json;odata=verbose"
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          req.Timeout <- Timeout.Infinite
          match items with
          | Files -> 
            let b = File.ReadAllBytes(path)
            req.ContentLength <- b.LongLength
            req.GetRequestStream().Write(b, 0, b.Length)
          | Folders -> req.ContentLength <- 0L
          use! rsp = req.AsyncGetResponse()
          return
            match (rsp :?> HttpWebResponse).StatusCode = HttpStatusCode.Created with
            | true -> None | false -> Some (path, spoUrls) } |> Async.Catch

  let rec internal copyHelper root path host urlRoot o365 digest items = 
    match items with
      | Folders -> 
        copyHelper root path host urlRoot o365 digest SPOitems.Files
        Directory.EnumerateDirectories(path)
      | Files -> Directory.EnumerateFiles(path)
    |> Ext.Seq.split 1000
    |> Seq.iter(fun xs ->
      xs
      |> Array.Parallel.map(fun x -> x, Util.nameUrl urlRoot root x)
      |> Array.Parallel.map(fun (x,(y,z)) -> x, Util.spoUrls host y z items)
      |> Array.Parallel.map(fun (x,y) -> CRUD.exists x y o365)
      |> Async.Parallel
      |> Async.RunSynchronously
      |> fun ys ->
        let success = 
          ys 
          |> Array.Parallel.choose(Either.chooseLeft)
          |> Array.Parallel.choose(id)

        success
        |> Array.Parallel.map(
          fun (x,y) -> CRUD.create x y o365 digest items)
        |> Async.Parallel
        |> Async.RunSynchronously
        |> fun zs ->
          let failure  = ys |> Array.Parallel.choose(Either.chooseRight)
          let failure' = zs |> Array.Parallel.choose(Either.chooseRight)

          failure  |> Array.Parallel.iter(log LogLevel.Error)
          failure' |> Array.Parallel.iter(log LogLevel.Error)

          match items with
            | Files -> ()
            | Folders ->
              zs 
              |> Array.Parallel.choose(Either.chooseLeft)
              |> Array.Parallel.choose(id)
              |> Array.Parallel.map(fun (x,_) -> x)
              |> Array.Parallel.iter(
                fun x -> copyHelper root x host urlRoot o365 digest items))
  
  let copy local (url:Uri) usr pwd = 
    let host = Uri(url.GetLeftPart(UriPartial.Authority))
    let urlRoot = url.ToString().Replace(host.ToString(),String.Empty)
    let o365 = Office365.getCookieContainer host usr pwd, Office365.userAgent
    let digest = SPO.CRUD.formDigest o365 host

    copyHelper local local (host.ToString()) urlRoot o365 digest SPOitems.Folders