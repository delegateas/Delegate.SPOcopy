namespace Delegate

module SPOcopy = 
  open System
  open System.IO
  open System.Net
  open System.Text.RegularExpressions
  open System.Threading
  open System.Xml.Linq

  open Microsoft.FSharp.Reflection

  type SPOitems = Folders | Files
  type FormDigest = { value : string; timeout : float; created : DateTime } with
    static member create v t d = { value = v; timeout = t; created = d }
    member x.soonExpire items = 
      let y = items |> function | Folders -> 5. | Files -> 10.
      DateTime.Now.AddSeconds(x.timeout / y) > x.created.AddSeconds (x.timeout)

  [<AutoOpen>]
  module Either = 
    type ('a,'b) Either = Choice<'a,'b>
    let success x : Either<'a,'b> = Choice1Of2 x
    let failure x : Either<'a,'b> = Choice2Of2 x

    let (|Success|Failure|) = function
      | Choice1Of2 x -> Success x | Choice2Of2 x -> Failure x

    let inline bind f    = function | Success a -> f a | Failure b -> failure b
    let inline (>>=) m f = bind f m

    let succeeded  = function | Success s -> Some s | Failure _ -> None
    let failed     = function | Success _ -> None   | Failure b -> Some b

  [<AutoOpen>]
  module Logger =
    // Info    = (1 <<< 0) ||| (1 <<< 2)
    // Warning = (1 <<< 1) ||| (1 <<< 2)
    // Error   = (1 <<< 2)
    // Verbose = (1 <<< 0) ||| (1 <<< 1) ||| (1 <<< 2)
    type LogLevel = 
      | Info    = 5 // Indicates logs for an informational message and errors
      | Warning = 6 // Indicates logs for a warning and errors
      | Error   = 4 // Indicates logs for an error.
      | Verbose = 7 // Indicates logs at all levels.

    let info    = LogLevel.Info
    let warning = LogLevel.Warning
    let error   = LogLevel.Error
    let verbose = LogLevel.Verbose

    let ts () = DateTime.Now.ToString("o") // ISO-8601
    let cw (s:string) = Console.WriteLine(s)
    let cew (s:string) = Console.Error.WriteLine(s)

    type Log = 
      { level : LogLevel }
      member x.print l y = 
        match x.level.HasFlag l with
        | false -> ()
        | true -> 
          let msg = sprintf "%s - %A: %A" (ts()) l y
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
      let formDigestHelper o365 (host:Uri) (log:Log) = 
        let ns = @"http://schemas.microsoft.com/ado/2007/08/dataservices"
        try 
          let created () = DateTime.Now
          let cookie, agent = o365
          let url = Uri(host.ToString() + "_api/contextinfo")
          let req = HttpWebRequest.Create(url) :?> HttpWebRequest
          req.Method <- "POST"
          req.ContentLength <- 0L
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          req.Timeout <- Timeout.Infinite
          req.ReadWriteTimeout <- Timeout.Infinite
          use rsp = req.GetResponse() :?> HttpWebResponse
          use stream = new StreamReader(rsp.GetResponseStream())
          let xml = stream.ReadToEnd()
          let xml' = XDocument.Parse(xml)
          let value = 
            xml'.Descendants()
            |> Seq.where (fun x -> x.Name = XNamespace.Get(ns) + "FormDigestValue")
            |> Seq.head
            |> fun x -> x.Value
          let timeout = 
            xml'.Descendants()
            |> Seq.where (fun x -> x.Name = XNamespace.Get(ns) + "FormDigestTimeoutSeconds")
            |> Seq.head
            |> fun x -> Double.Parse(s = x.Value)
          FormDigest.create value timeout (created ())
        with ex -> log.print error ex; raise ex

      let formDigest o365 (host:Uri) (digest:FormDigest) items log = 
        match digest.soonExpire items with
        | true ->  formDigestHelper o365 host log
        | false -> digest

      let exists (path:string) spoUrls o365 = async{
          let url, _ = spoUrls
          let cookie, agent = o365
          let req = HttpWebRequest.Create(requestUriString = url) :?> HttpWebRequest
          req.Method <- "GET"
          req.Accept <- "application/json;odata=verbose"
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          req.Timeout <- Timeout.Infinite
          req.ReadWriteTimeout <- Timeout.Infinite
          use! rsp = req.AsyncGetResponse()
          use stream = rsp.GetResponseStream()
          use reader = new StreamReader(stream)
          let json = reader.ReadToEnd()
          return 
            match json.Contains("\"Exists\": true") with
            | true -> None | false -> Some (path, spoUrls) } |> Async.Catch

      let create (path:string) spoUrls o365 digest items = async{
          let _, url = spoUrls
          let cookie, agent = o365
          let req = HttpWebRequest.Create(requestUriString = url) :?> HttpWebRequest
          req.Headers.Add("X-RequestDigest", digest().value)
          req.Method <- "POST"
          req.Accept <- "application/json;odata=verbose"
          req.CookieContainer <- cookie
          req.UserAgent <- agent
          req.Timeout <- Timeout.Infinite
          req.ReadWriteTimeout <- Timeout.Infinite
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

  let rec internal copyHelper root path host urlRoot o365 digest items (log:Log) = 
    let host' = host.ToString()
    match items with
      | Folders ->
        copyHelper root path host urlRoot o365
          (SPO.CRUD.formDigest o365 host digest items log) SPOitems.Files log
        Directory.EnumerateDirectories(path)
      | Files -> Directory.EnumerateFiles(path)
    |> Seq.split 1000
    |> Seq.iter(fun xs ->
      log.print verbose
        (sprintf "Creating/Uploading %4i %s, path: %s"
          (Array.length xs) (Util.ducToString items) path)
  
      xs
      |> Array.Parallel.map(fun x -> x, Util.nameUrl urlRoot root x)
      |> Array.Parallel.map(fun (x,(y,z)) -> x, Util.spoUrls host' y z items)
      |> Array.Parallel.map(fun (x,y) -> CRUD.exists x y o365)
      |> Async.Parallel
      |> Async.RunSynchronously
      |> fun ys ->
        let success = 
          ys 
          |> Array.Parallel.choose(Either.succeeded)
          |> Array.Parallel.choose(id)

        success
        |> Array.Parallel.map(fun (x,y) -> 
            CRUD.create x y o365 
              (fun () -> SPO.CRUD.formDigest o365 host digest items log) items)
        |> Async.Parallel
        |> Async.RunSynchronously
        |> fun zs ->
          let failure  = ys |> Array.Parallel.choose(Either.failed)
          let failure' = zs |> Array.Parallel.choose(Either.failed)

          failure  |> Array.iter(log.print error)
          failure' |> Array.iter(log.print error)

          match items with
            | Files -> ()
            | Folders ->
              zs 
              |> Array.Parallel.choose(Either.succeeded)
              |> Array.Parallel.choose(id)
              |> Array.Parallel.map(fun (x,_) -> x)
              |> Array.Parallel.iter(
                fun x ->
                  copyHelper root x host urlRoot o365
                    (SPO.CRUD.formDigest o365 host digest items log) items log))
  
  let copy local (url:Uri) usr pwd loglvl = 
    let host = Uri(url.GetLeftPart(UriPartial.Authority))
    let urlRoot = url.ToString().Replace(host.ToString(),String.Empty)
    let o365 = Office365.getCookieContainer host usr pwd, Office365.userAgent
    let log = { Log.level = loglvl }

    log.print info "SharePoint Online copy (SPOcopy) - Started"

    copyHelper local local host urlRoot o365
      (SPO.CRUD.formDigestHelper o365 host log) SPOitems.Folders log

    log.print info "SharePoint Online copy (SPOcopy) - Finished"
