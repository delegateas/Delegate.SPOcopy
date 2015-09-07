namespace Delegate

module Office365 = 
  open System
  open System.IO
  open System.Net
  open System.Net.Security
  open System.ServiceModel
  open System.ServiceModel.Channels
  open System.Text
  open System.Xml
  open System.Xml.Linq
  open Microsoft.IdentityModel.Protocols.WSTrust
  
  [<Literal>]
  let userAgent = 
    "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
  
  type SamlTokenInfo = 
    { FedAuth : string
      rtFa : string
      Expires : DateTime
      Host : Uri }
  
  [<ServiceContract>]
  type IWSTrustFeb2005Contract = 
    [<OperationContract(
      ProtectionLevel = ProtectionLevel.EncryptAndSign, 
      Action = "http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue", 
      ReplyAction = "http://schemas.xmlsoap.org/ws/2005/02/trust/RSTR/Issue", 
      AsyncPattern = true)>]
    abstract BeginIssue : request:Message * callback:AsyncCallback * state:obj
     -> IAsyncResult
    
    abstract EndIssue : result:IAsyncResult -> Message
  
  type WSTrustFeb2005ContractClient(binding, remoteAddress) = 
    inherit ClientBase<IWSTrustFeb2005Contract>((binding : Binding), 
                                                remoteAddress)
    interface IWSTrustFeb2005Contract with
      member __.BeginIssue(request : Message, callback : AsyncCallback, 
                           state : obj) : IAsyncResult = 
        base.Channel.BeginIssue(request, callback, state)
      member __.EndIssue(asyncResult) = base.Channel.EndIssue(asyncResult)
  
  type RequestBodyWriter(serializer : WSTrustRequestSerializer, rst : RequestSecurityToken) = 
    inherit BodyWriter(false)
    override __.OnWriteBodyContents(writer : XmlDictionaryWriter) = 
      serializer.WriteXml(rst, writer, new WSTrustSerializationContext())
  
  let private createRequest (url : Uri) = 
    let r = HttpWebRequest.Create(url) :?> HttpWebRequest
    r.Method <- "POST"
    r.ContentType <- "application/x-www-form-urlencoded"
    r.CookieContainer <- CookieContainer()
    r.AllowAutoRedirect <- false
    r.UserAgent <- userAgent
    r
  
  let private getStsResponse (sts : Uri) (realm : Uri) username password = 
    let rst = 
      RequestSecurityToken
        (RequestType = WSTrustFeb2005Constants.RequestTypes.Issue, 
         AppliesTo = EndpointAddress(realm.ToString()), 
         KeyType = WSTrustFeb2005Constants.KeyTypes.Bearer, 
         TokenType = Microsoft.IdentityModel.Tokens.SecurityTokenTypes.Saml11TokenProfile11)
    let trustSerializer = WSTrustFeb2005RequestSerializer()
    let b = WSHttpBinding()
    b.Security.Mode <- SecurityMode.TransportWithMessageCredential
    b.Security.Message.ClientCredentialType <- MessageCredentialType.UserName
    b.Security.Message.EstablishSecurityContext <- false
    b.Security.Message.NegotiateServiceCredential <- false
    b.Security.Transport.ClientCredentialType <- HttpClientCredentialType.None
    let a = EndpointAddress(sts.ToString())
    use trustClient = new WSTrustFeb2005ContractClient(b, a)
    trustClient.ClientCredentials.UserName.UserName <- username
    trustClient.ClientCredentials.UserName.Password <- password
    let response = 
      (trustClient :> IWSTrustFeb2005Contract)
        .EndIssue((trustClient :> IWSTrustFeb2005Contract)
          .BeginIssue(Message.CreateMessage
                        (MessageVersion.Default, 
                         WSTrustFeb2005Constants.Actions.Issue, 
                         RequestBodyWriter(trustSerializer, rst)), null, null))
    trustClient.Close()
    use reader = response.GetReaderAtBodyContents()
    reader.ReadOuterXml()
  
  let getSamlToken (host : Uri) username password = 
    let office365STS = Uri "https://login.microsoftonline.com/extSTS.srf"
    let wsse = 
      "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"
    let wsu = 
      "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"
    let wreply = 
      Uri
        (host.GetLeftPart(UriPartial.Authority) 
         + "/_forms/default.aspx?wa=wsignin1.0")
    let stsResponse = getStsResponse office365STS wreply username password
    let xml = XDocument.Parse(stsResponse)
    
    let securityToken = 
      xml.Descendants()
      |> Seq.where (fun r -> r.Name = XName.Get("BinarySecurityToken", wsse))
      |> Seq.map (fun r -> r.Value)
      |> Seq.head
    
    let tokenExpiration = 
      xml.Descendants()
      |> Seq.where (fun r -> r.Name = XName.Get("Expires", wsu))
      |> Seq.map (fun r -> Convert.ToDateTime(r.Value))
      |> Seq.head
    
    let request = createRequest (wreply)
    let data = Encoding.UTF8.GetBytes(securityToken)
    use stream = request.GetRequestStream()
    stream.Write(data, 0, data.Length)
    use response = request.GetResponse() :?> HttpWebResponse
    // handle redirect added May 2011 for P-subscriptions
    if response.StatusCode = HttpStatusCode.MovedPermanently then 
      let request2 = createRequest (Uri(response.Headers.["Location"]))
      use stream2 = request2.GetRequestStream()
      stream2.Write(data, 0, data.Length)
      use response2 = request2.GetResponse() :?> HttpWebResponse
      { FedAuth = response2.Cookies.["FedAuth"].Value
        rtFa = response2.Cookies.["rtFa"].Value
        Expires = tokenExpiration
        Host = request2.RequestUri }
    else 
      { FedAuth = response.Cookies.["FedAuth"].Value
        rtFa = response.Cookies.["rtFa"].Value
        Expires = tokenExpiration
        Host = request.RequestUri }
  
  let getCookieContainer (host : Uri) (username : string) (password : string) : CookieContainer = 
    let samlTokenInfo = getSamlToken host username password
    let cookieContainer = CookieContainer()
    // add FedAuth cookie
    let samlAuth = 
      Cookie
        ("FedAuth", samlTokenInfo.FedAuth, Expires = samlTokenInfo.Expires, 
         Path = "/", Secure = (samlTokenInfo.Host.Scheme = "https"), 
         HttpOnly = true, Domain = samlTokenInfo.Host.Host)
    cookieContainer.Add(samlAuth)
    // add rtFA (sign-out) cookie added March 2011
    let rtFa = 
      Cookie
        ("rtFA", samlTokenInfo.rtFa, Expires = samlTokenInfo.Expires, Path = "/", 
         Secure = (samlTokenInfo.Host.Scheme = "https"), HttpOnly = true, 
         Domain = samlTokenInfo.Host.Host)
    cookieContainer.Add(rtFa)
    cookieContainer