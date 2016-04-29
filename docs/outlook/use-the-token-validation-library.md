
# Verwenden der Exchange-Tokenüberprüfungsbibliothek
Verwenden der EWS Managed API-Überprüfungsbibliothek zum Überprüfen eines Exchange-Identitätstokens

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Sie können die Clients des Outlook-Add-Ins mithilfe eines Identitätstokens identifizieren, die Ihr Add-in von einem Server mit Exchange Server 2013 anfordert. Das Token, das als JSON-Webtoken formatiert ist, dient für ein E-Mail-Konto auf einem Exchange-Server als eindeutiger Bezeichner. Die Exchange Web Services (EWS) Managed API stellt Hilfsklassen bereit, um die Verwendung des Identitätstokens zu vereinfachen.

## Voraussetzungen für die Verwendung der Überprüfungsbibliothek
<a name="bk_prerequisites"> </a>

Zum Überprüfen eines Exchange-Identitätstokens benötigen Sie die EWS Managed API-Authentifizierungsbibliothek und die Windows Identity Foundation (WIF) in Verbindung mit einer DLL, mit der die WIF um Handler für JSON-Token erweitert wird. Stellen Sie sicher, dass Sie die folgenden Ressourcen herunterladen:


- [Managed API für Exchange-Webdienste](http://go.microsoft.com/fwlink/?LinkID=255472)
    
- [Windows Identity Foundation ](http://www.microsoft.com/de-de/download/details.aspx?id=17331)
    
- [Windows.IdentityModel.Extensions.dll für 32-Bit-Anwendungen](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll für 64-Bit-Anwendungen](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Überprüfen des Exchange-Identitätstokens
<a name="bk_validate"> </a>

Die EWS Managed API-Überprüfungsbibliothek stellt die  **AppIdentityToken** -Klasse zum Verwalten der Exchange-Identitätstoken bereit. Die folgende Methode verdeutlicht die Erstellung einer **AppIdentityToken** -Instanz und den Aufruf der **Validate** -Methode zum Überprüfen der Gültigkeit des Tokens.


```
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

        private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
        {
            try
            {
                AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
                token.Validate(new Uri(hostUri));

                return token;
            }
            catch (TokenValidationException ex)
            {
                throw new ApplicationException("A client identity token validation error occurred.", ex);
            }
        }

```


## Zusätzliche Ressourcen
<a name="bk_additionalresources"> </a>


- [Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken](c0520a1e-d9ba-495a-a99f-6816d7d2a23e.md)
    
- [Innenleben des Exchange-Identitätstokens](e4085245-73ae-4480-98ac-15593daa5fd7.md)
    
- [Überprüfen eines Exchange-Identitätstokens](8503a3e8-458a-4a4e-9e95-65cd7bb1954d.md)
    
