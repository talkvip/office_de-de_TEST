
# Authentifizieren eines Benutzers mit einem Identitätstoken für Exchange
Sehen Sie sich ein Codebeispiel an, in dem gezeigt wird, wie ein Authentifizierungsobjekt der eindeutigen Identität entspricht, die durch ein Identitätstoken mit einem Satz von Anmeldeinformationen für einen Dienst dargestellt wird.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Sie können ein Authentifizierungsschema mit einmaligem Anmelden (Single Sign-On, SSO) für einen Informationsdienst verwenden, der es Kunden, die Outlook-Add-Ins nutzen, ermöglicht, über ihre Exchange-Server-Anmeldeinformationen eine Verbindung mit Ihrem Dienst herzustellen. In diesem Artikel wird der Abgleich von Anmeldeinformationen mithilfe eines einfachen, auf einem  **Dictionary** -Objekt basierenden Benutzerdatenspeichers erläutert.

(../../images  Hierbei handelt es sich um ein einfaches Beispiel von SSO, das nicht für die Verwendung in Ihrem Produktionscode gedacht ist. Des Weiteren müssen Sie – wie immer, wenn es um Identität und Authentifizierung geht – sicherstellen, dass Ihr Code die Sicherheitsanforderungen Ihrer Organisation erfüllt.


## Voraussetzungen für die Verwendung der SSO-Authentifizierung


Um ein Identitätstoken für SSO verwenden zu können, muss Ihre Dienstanwendung über ein gültiges Identitätstoken verfügen. Weitere Informationen zu Identitätstoken und deren Anforderung und Überprüfung finden Sie in den folgenden Artikeln:


- [Innenleben des Exchange-Identitätstokens](../outlook/inside-the-identity-token.md)
    
- [Aufrufen eines Diensts von einem Outlook-Add-In mithilfe eines Identitätstokens in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Verwenden der Exchange-Tokenüberprüfungsbibliothek](../outlook/use-the-token-validation-library.md), wenn Sie verwalteten Code verwenden, oder [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md), wenn Sie Ihre eigene Methode für die Überprüfung von Token erstellen.
    

## Authentifizieren eines Benutzers


Im folgenden Codebeispiel wird ein einfaches Authentifizierungsobjekt gezeigt, das der eindeutigen Identität entspricht, die durch ein Identitätstoken mit einem Satz von Anmeldeinformationen für einen Dienst dargestellt ist. Die Klasse  **TokenAuthentication** bietet die Methode **GetResponseFromService**, die eine Antwort für bereits authentifizierte Token zurückgibt oder den Benutzer zur Eingabe von Anmeldeinformationen auffordert, die authentifiziert und dem Identitätstoken zugeordnet werden können. Der Code ist nicht vollständig; es wird davon ausgegangen, dass Sie die folgenden Objekte und Methoden bereitstellen:



|**Objekt/Methode**|**Beschreibung**|
|:-----|:-----|
|**LocalCredentials** -Objekt|Stellt die Anmeldeinformationen des Benutzers für Ihren Dienst dar. Die Struktur des Objekts hängt von den Anforderungen Ihres Diensts ab.|
|**IdentityToken** -Objekt|Enthält ein Benutzeridentitätstoken, das von einem Outlook-Add-In an Ihren Dienst gesendet wurde. Das Objekt muss mindestens den eindeutigen Exchange-Bezeichner des Benutzers und die Authentifizierungsmetadaten-URL für den Server enthalten, von dem das Token ausgegeben wurde. In diesem Beispiel wird das im Artikel [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md) definierte Identitätstokenobjekt verwendet.|
|**JsonResponse** -Objekt|Stellt die Antwort von Ihrem Dienst dar. Das Objekt kann in ein JSON-Objekt serialisiert werden.|
|**CallService** -Methode|Führt einen Aufruf an Ihren Dienst mit einem  **LocalCredentials** -Objekt durch, das die Anmeldeinformationen des Benutzers für den Dienst enthält und einem Objekt, das Daten für die Dienstanforderung enthält. Sind die Anmeldeinformationen gültig, gibt diese Methode ein **JsonReponse** -Objekt zurück, das die Ergebnisse der Anforderung enthält. Sind die Anmeldeinformationen ungültig, gibt diese Methode " **null**" zurück.|
|**GetCredentialsResponse** -Methode|Gibt ein  **JsonReponse** -Objekt zurück, das Ihr Mail-Office-Add-in als Anforderung für die Anmeldeinformationen für den Dienst erkennt.|
|**LocalCredentialsAreValid** -Methode|Gibt " **true**" zurück, wenn die dem Dienst bereitgestellten Anmeldeinformationen gültig sind. Andernfalls wird " **false**" zurückgegeben.|

(../../images  Hierbei handelt es sich lediglich um einen Vorschlag für die Verwendung des Identitätstokens. Wie immer, wenn es um Identität und Authentifizierung geht, müssen Sie sicherstellen, dass Ihr Code die Sicherheitsanforderungen Ihrer Organisation erfüllt.


```
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## Authentifizieren eines Benutzers mit der verwalteten Überprüfungsbibliothek


Wenn Sie die verwaltete Bibliothek für die Überprüfung von Identitätstoken verwenden, müssen Sie den eindeutigen Schlüssel nicht berechnen. Die  **UniqueUserIdentification** -Eigenschaft der **AppIdentityToken** -Klasse kann direkt als eindeutiger Schlüssel für den Benutzer verwendet werden. Im folgenden Codebeispiel werden die Änderungen an der **GetResponseFromService** -Methode aus dem vorherigen Beispiel erläutert, die Sie vornehmen müssen, um die **AppIdentityToken** -Klasse verwenden zu können.


```
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## Zusätzliche Ressourcen



- [Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken](../outlook/authentication.md)
    
- [Aufrufen eines Diensts von einem Outlook-Add-In mithilfe eines Identitätstokens in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Verwenden der Exchange-Tokenüberprüfungsbibliothek](../outlook/use-the-token-validation-library.md)
    
- [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md)
    
