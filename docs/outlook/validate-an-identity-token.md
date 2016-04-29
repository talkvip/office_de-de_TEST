
# Überprüfen eines Exchange-Identitätstokens
In diesem Thema erfahren Sie, wie Sie das Exchange 2013-Identitätstoken überprüfen, das die E-Mail-Konten Ihres Outlook-Add-In-Kunden mit den von Ihrem Webdienst bereitgestellten Informationen verknüpft.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


![Verwandte Codeausschnitte und Beispiel-Apps](../../images/mod_icon_links_samples.png)[Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)

Ihr Outlook-Add-In kann Ihnen ein Identitätstoken senden. Bevor Sie der Anforderung vertrauen, müssen Sie jedoch das Token überprüfen, um sicherzustellen, dass es auch wirklich von dem Exchange-Server stammt. Anhand der Beispiele in diesem Artikel wird gezeigt, wie Sie das Exchange-Identitätstoken mithilfe eines in C# geschriebenen Validierungsobjekts überprüfen. Für die Überprüfung können Sie jedoch eine beliebige Programmiersprache verwenden. Die zum Überprüfen des Tokens erforderlichen Schritte sind im [JSON Web Token (JWT)-Internetentwurf](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl) beschrieben.

Es wird empfohlen, einen aus vier Schritten bestehenden Vorgang zum Überprüfen des Identitätstokens und Abrufen des eindeutigen Bezeichners des Benutzers zu verwenden. Extrahieren Sie zunächst das JSON Web Token (JWT) aus einer base64-URL-codierten Zeichenfolge. Stellen Sie anschließend sicher, dass das Token wohlgeformt ist, dass es für Ihr Outlook-Add-In bestimmt und nicht abgelaufen ist und dass Sie eine gültige URL für das Dokument mit den Authentifizierungsmetadaten extrahieren können. Rufen Sie als Nächstes das Dokument mit den Authentifizierungsmetadaten vom Exchange-Server ab, und überprüfen Sie die an das Identitätstoken angefügte Signatur. Berechnen Sie schließlich einen eindeutigen Bezeichner für den Benutzer, indem Sie einen Hash für die Exchange-ID des Benutzers mit der URL des Authentifizierungsmetadaten-Dokuments erstellen. Dieser Vorgang scheint insgesamt komplex zu sein, die einzelnen Schritte sind jedoch recht einfach.
Sie können die Lösung, die diese Beispiele enthält, unter [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken) herunterladen.
 **Inhalt dieses Artikels**

![Eins](../../images/mod_icon_one.png)  [Einrichtung zum Überprüfen Ihres Identitätstokens](#einrichtung-zum-überprüfen-ihres-identitätstokens)

![Zwei](../../images/mod_icon_two.png)  [Extrahieren des JSON Web Token](#extrahieren-des-json-web-token)

![Drei](../../images/mod_icon_three.png)  [Analysieren des JWT](#analysieren-des-jwt)

![Vier](../../images/mod_icon_four.png)  [Überprüfen der Identitätstokensignatur](#überprüfen-der-identitätstokensignatur)

![fünf](../../images/mod_icon_five.png)  [Berechnen der eindeutigen ID für ein Exchange-Konto](#berechnen-der-eindeutigen-id-für-ein-exchange-konto)

![Verwandte Codeausschnitte und Beispiel-Apps](../../images/mod_icon_links_samples.png)  [Hilfsprogrammobjekte](#hilfsprogrammobjekte)


## Einrichtung zum Überprüfen Ihres Identitätstokens
<a name="setup"> </a>

Die Codebeispiele in diesem Artikel basieren auf der Windows Identity Foundation (WIF) sowie einer DLL, durch die die WIF mit Handlern für JSON-Tokens erweitert werden. Die erforderlichen Assemblys können von den folgenden Speicherorten herunterladen:


- [Windows Identity Foundation](http://msdn.microsoft.com/de-de/security/aa570351)
    
- [Windows.IdentityModel.Extensions.dll für 32-Bit-Anwendungen](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll für 64-Bit-Anwendungen](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Extrahieren des JSON Web Token
<a name="extractthejsonwebtoken"> </a>

Mit der  **Decode** -Factorymethode wird das JWT auf dem Exchange-Server in die drei Zeichenfolgen unterteilt, aus denen das Token besteht. Anschließend werden mithilfe der **Base64Decode** -Methode (siehe zweites Beispiel) der JWT-Header und die Last in JSON-Zeichenfolgen decodiert. Die Zeichenfolgen werden an den **JsonToken** -Konstruktor übergeben, wo der Inhalt des JWT überprüft wird und eine neue **JsonToken** -Objektinstanz zurückgegeben wird.


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

Mit der  **Base64Decode** -Methode wird die Decodierungslogik implementiert, die im[JSON Web Token (JWT)-Internetentwurf](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl) im Anhang mit den Hinweisen zur Implementierung von Base64-URL-Codierung ohne Auffüllung beschrieben wird.




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## Analysieren des JWT
<a name="parsethejwt"> </a>

Der Konstruktor für das  **JsonToken** -Objekt prüft die Struktur und den Inhalt des JWT und bestimmt, ob es gültig ist. Es empfiehlt sich, diesen Vorgang auszuführen, bevor Sie das Dokument mit den Authentifizierungsmetadaten anfordern. Falls das JWT nicht die ordnungsgemäßen Ansprüche enthält oder falls dessen Lebensdauer abgelaufen ist, können Sie einen Aufruf des Exchange-Servers und die damit verbundene Verzögerung vermeiden.

Der Konstruktor ruft Hilfsmethoden auf, um zu bestimmen, ob die unterschiedlichen Ansprüche vorhanden sind und sich innerhalb des gültigen Bereichs befinden. Falls ein Problem vorliegt, löst die Hilfsmethode einen Ausnahmefehler für die Anwendung aus. Falls keine Ausnahmefehler ausgelöst werden, wird die  **IsValid** -Eigenschaft auf **true** festgelegt, und das Token ist bereit für die Signaturüberprüfung.

Die verschiedenen Hilfsmethoden werden weiter unten in diesem Artikel näher beschrieben.




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### ValidateHeader-Methode

Mit der  **ValidateHeader** -Methode wird sichergestellt, dass die erforderlichen Ansprüche im Tokenheader vorhanden sind und dass die Ansprüche die richtigen Werte aufweisen. Der Header muss wie im Folgenden beschrieben definiert werden. Andernfalls löst die Methode einen Ausnahmefehler aus und der Vorgang wird beendet.

{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<Fingerabdruck>" }
```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### ValidateLifetime-Methode

Zwei Datumsangaben werden im JWT angegeben: "nbf" ("not before") entspricht dem Datum und der Uhrzeit, an dem bzw. zu der das Token gültig wird. "exp" entspricht der Uhrzeit, zu der das Token abläuft. Nur Token zwischen diesen beiden Datumsangaben sollten als gültig betrachtet werden. Um kleinere Unterschiede zwischen den Uhrzeiteinstellungen des Servers und des Clients abzufangen, werden von dieser Methode Token bis zu fünf Minuten vor und nach den im Token angegebenen Zeiten validiert.


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

Die Datumsangaben  **validFrom** ("nbf") und **validTo** ("exp") werden seit dem 1. Januar 1970, dem Beginn der Unix-Epoche, als die Anzahl von Sekunden gesendet. Die Datums- und Uhrzeitangaben werden mithilfe der koordinierten Weltzeit (UTC) berechnet, um Probleme durch unterschiedliche Zeitzonen zwischen dem Exchange-Server und dem Server, auf dem der Überprüfungscode ausgeführt wird, zu vermeiden.


### ValidateAudience-Methode

Das Identitätstoken ist nur gültig für die Add-In, von der es angefordert wurde. Mit der Methode  **ValidateAudience** wird der Zielgruppenanspruch im Token überprüft, um sicherzustellen, das er mit der erwarteten URL für das Outlook-Add-In übereinstimmt.


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### ValidateVersion-Methode

Mit der  **ValidateVersion** -Methode wird die Version des Identitätstokens überprüft und sichergestellt, dass sie der erwarteten Version entspricht. Unterschiedliche Versionen des Tokens können unterschiedliche Ansprüche aufweisen. Durch das Überprüfen der Version wird sichergestellt, dass die erwarteten Ansprüche im Identitätstoken vorhanden sind.


```
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### ValidateMetadataLocation-Methode

Das auf dem Exchange-Server gespeicherte Authentifizierungsmetadatenobjekt enthält die Informationen, die zum Überprüfen der Signatur im Identitätstoken erforderlich sind. Mit der  **ValidateMetadataLocation** -Methode wird sichergestellt, dass eine Authentifizierungsmetadaten-URL im Identitätstoken vorhanden ist, wobei im nächsten Schritt überprüft wird, ob eine Signatur vorhanden ist.


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## Überprüfen der Identitätstokensignatur
<a name="validatetheidentitytokensignature"> </a>

Nachdem Sie nun wissen, dass das JWT die erforderlichen Ansprüche zum Überprüfen der Signatur enthält, können Sie mithilfe der Windows Identity Foundation (WIF) und den WIF-Erweiterungen die Signatur im Token überprüfen. Zum Überprüfen der Signatur benötigen Sie die folgenden Information:


- Die ursprüngliche Base64-URL-codierte Identitätstokenzeichenfolge, die vom Exchange-Server gesendet wurde
    
- Den Speicherort des Dokuments mit den Authentifizierungsmetadaten vom JWT
    
- Die Zielgruppen-URL vom JWT
    
In diesem Beispiel ruft der Konstruktor für ein  **IdentityToken** -Objekt das Dokument mit den Authentifizierungsmetadaten vom Exchange-Server ab und überprüft die Signatur im Identitätstoken. Falls das Identitätstoken gültig ist, können Sie mit der **IdentityToken** -Objektinstanz die im Identitätstoken enthaltene eindeutige Benutzer-ID abrufen.




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

Mit dem Großteil des Codes im  **IdentityToken** -Objektkonstruktor werden die Eigenschaften in der Instanz mit den Ansprüchen vom Exchange-Server festgelegt. Der Konstruktor ruft die **GetSecurityTokenHandler** -Methode auf, um einen Tokenhandler abzurufen, mit dem das Exchange-Identitätstoken überprüft wird. Mit der **GetSecurityTokenHandler** -Methode werden die beiden Hilfsmethoden **GetMetadataDocument** und **GetSigningCertificate** aufgerufen, mit denen das Signaturzertifikat vom Exchange-Server abgerufen wird. Diese Methoden werden in den folgenden Abschnitten beschrieben.


### GetSecurityTokenHandler-Methode

Mit der  **GetSecurityTokenHandler** -Methode wird ein WIF-Tokenhandler zurückgegeben, mit dem das Identitätstoken überprüft wird. Mit dem Großteil des Codes dieser Methode wird der Tokenhandler initialisiert, um die Überprüfung vorzunehmen. Mit dieser Methode wird jedoch auch die **GetSigningCertificate** -Methode aufgerufen, um das X.509-Zertifikat, mit dem das Token signiert wird, vom Exchange-Server abzurufen.


```
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### GetSigningCertificate-Methode

Mit der  **GetSigningCertificate** -Methode wird die **GetMetadataDocument** -Methode aufgerufen, um die Authentifizierungsmetadaten vom Exchange-Server abzurufen. Anschließend wird das erste X.509-Zertifikat im Dokument mit den Authentifizierungsmetadaten zurückgegeben. Falls das Dokument vorhanden ist, wird ein Ausnahmefehler für die Anwendung ausgelöst.


```
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### GetMetadataDocument-Methode

Das Dokument mit den Authentifizierungsmetadaten enthält die Informationen, die Sie zum Überprüfen der Signatur im Exchange-Identitätstoken benötigen. Das Dokument wird als JSON-Zeichenfolge gesendet. Mit der  **GetMetatDataDocument** -Methode wird das Dokument aus dem im Exchange-Identitätstoken angegebenen Speicherort angefordert. Es wird dann ein Objekt zurückgegeben, das die JSON-Zeichenfolge als Objekt enthält. Falls die URL kein Dokument mit den Authentifizierungsmetadaten enthält, wird ein Ausnahmefehler für die Anwendung ausgelöst.


```
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

Der Exchange-Server verwendet standardmäßig ein selbstsigniertes X.509-Zertifikat zum Authentifizieren von Anforderungen für das Dokument mit den Authentifizierungsmetadaten. Sofern Sie kein Zertifikat installieren, das bis zu einem Stammserver nachverfolgt, müssen Sie eine Rückrufmethode für die Zertifikatüberprüfung erstellen, da andernfalls die Anforderung für das Dokument mit den Authentifizierungsmetadaten fehlschlägt. 

Mit der  **ServicePointManager** -Klasse im .NET Framework-Namespace **System.Net** können Sie eine Rückrufmethode für die Überprüfung einbinden, indem Sie die **ServerCertificateValidationCallback** -Eigenschaft festlegen. Ein Beispiel für eine Rückrufmethode für die Zertifikatüberprüfung, die zu Entwicklungs- und Testzwecken geeignet ist, finden Sie im Artikel[Überprüfen von X509-Zertifikaten](http://msdn.microsoft.com/de-de/library/dd633677%28EXCHG.80%29.aspx).


 **Sicherheitshinweis**  Falls Sie eine Rückrufmethode für die Zertifikatüberprüfung verwenden, müssen Sie sicherstellen, dass sie die Sicherheitsanforderungen Ihres Unternehmens erfüllt.


## Berechnen der eindeutigen ID für ein Exchange-Konto
<a name="uniqueid"> </a>

Sie können einen eindeutigen Bezeichner für ein Exchange-Konto erstellen, indem Sie einen Hash für die URL des Authentifizierungsmetadaten-Dokuments mit dem Exchange-Bezeichner des Kontos erstellen. Wenn Sie über diesen eindeutigen Bezeichner verfügen, können Sie ihn verwenden, um ein System für einmaliges Anmelden (Single Sign-On, SSO) für Ihr Outlook-Add-In beim Webdienst zu erstellen. Weitere Informationen zur Verwendung des eindeutigen Bezeichners für SSO finden Sie unter [Authentifizieren eines Benutzers mit einem Identitätstoken für Exchange](bb28ca39-1780-4162-a899-7be5825beb8e.md)

Von der  **UniqueUserIdentification** -Eigenschaft wird mithilfe des Standard-SHA256-Anbieters aus dem **System.Security.Cryptography** -Namespace ein SHA256-Hash mit Salt der Exchange-ID und der Authentifizierungsmetadaten-URL erstellt.


 **Sicherheitshinweis**  Sie müssen einen Hash des Authentifizierungsmetadaten-Dokuments mit der Exchange-ID erstellen, um einen eindeutigen Bezeichner für das Konto zu erstellen. Wenn Sie nur die Exchange-ID verwenden, ist der Dienst möglicherweise für nicht autorisierte Benutzer zugänglich. Des Weiteren müssen Sie – wie immer, wenn es um Authentifizierung und Sicherheit geht – sicherstellen, dass bei Verwendung des mit dieser Methode erstellten eindeutigen Bezeichners die Sicherheitsanforderungen Ihrer Anwendung erfüllt werden.




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## Hilfsprogrammobjekte
<a name="utilityobjects"> </a>

Die Codebeispiele in diesem Artikel basieren auf einigen Hilfsprogrammobjekten, die Anzeigenamen für die verwendeten Konstanten liefern. In der folgenden Tabelle sind die Hilfsprogrammobjekte aufgelistet.


**Tabelle 1: Hilfsprogrammobjekte**


|**Objekt**|**Beschreibung**|
|:-----|:-----|
|**AuthClaimsType**|Erfasst die die vom Tokenüberprüfungscode verwendeten Anspruchsbezeichner in einem einzigen Speicherort.|
|**Config**|Stellt die Konstanten zum Überprüfen des Identitätstokens bereit. |
|**JsonAuthMetadataDocument**|Kapselt das vom Exchange-Server gesendete JSON-Dokument mit den Authentifizierungsmetadaten.|

### AuthClaimTypes-Objekt

Mit dem  **AuthClaimTypes** -Objekt werden die Anspruchsbezeichner, die vom Tokenüberprüfungscode verwendet werden, in einem einzigen Speicherort erfasst. Dieses Objekt enthält sowohl standardmäßige JWT-Ansprüche als auch die speziellen Ansprüche im Exchange-Identitätstoken.


```
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### Config-Objekt

Das  **Config** -Objekt enthält die Konstanten, mit denen das Identitätstoken überprüft wird, sowie eine Rückrufmethode für die Zertifikatüberprüfung, die Sie verwenden können, falls Ihr Server kein X509-Zertifikat aufweist, das bis zu einem Stammserver nachverfolgt.


 **Sicherheitshinweis**  Die Rückrufmethode für das Sicherheitszertifikat ist nur erforderlich, wenn Ihr Server das standardmäßige selbstsignierte Zertifikat verwendet. Mit der Rückrufmethode in diesem Beispiel wird  **false** zurückgegeben, wenn das Zertifikat selbstsigniert ist, weshalb sie durch eine Rückrufmethode ersetzt werden muss, welche die Sicherheitsanforderungen Ihres Unternehmens erfüllt. Ein Beispiel für eine Rückrufmethode für die Zertifikatüberprüfung, die zu Entwicklungs- und Testzwecken geeignet ist, finden Sie unter[Überprüfen von X509-Zertifikaten](http://msdn.microsoft.com/de-de/library/dd633677%28EXCHG.80%29.aspx).


```
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### JsonAuthMetadataDocument-Objekt

Mit dem  **JsonAuthMetadataDocument** -Objekt wird der Inhalt des Dokuments mit den Authentifizierungsmetadaten über Eigenschaften verfügbar gemacht.


```
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string[] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## Weitere Ressourcen
<a name="utilityobjects"> </a>


- [Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken](c0520a1e-d9ba-495a-a99f-6816d7d2a23e.md)
    
- [Innenleben des Exchange-Identitätstokens](e4085245-73ae-4480-98ac-15593daa5fd7.md)
    
