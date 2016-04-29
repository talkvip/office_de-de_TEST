
# Aufrufen eines Diensts von einem Outlook-Add-In mithilfe eines Identitätstokens in Exchange
Informationen zum Verwenden von Identitätstoken zum Verknüpfen der E-Mail-Konten Ihrer Outlook-Add-In-Kunden mit den Informationen, die Ihr Webdienst bereitstellt.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Ein Identitätstoken stellt eine eindeutige ID für alle Ihre Kunden bereit, mit deren Hilfe Sie den Dienst, den Sie anbieten, personalisieren können. Ihr Code kann beim Exchange-Server ein Identitätstoken mithilfe eines asynchronen Methodenaufrufs anfragen, der eine Zeichenfolge an Ihr Outlook-Add-In zurückgibt. Die Zeichenfolge enthält das JWT-Identitätstoken (JSON Web Token). Ihr Add-in muss das Token nicht entpacken, sondern kann es an so Ihren Webdienst übergeben, dass Ihr Dienst die Anforderung des Add-Ins authentifizieren kann.

Der Webdienst, der das Add-In unterstützt, muss auf dem Server ausgeführt werden, der auch die HTML- und JavaScript-Quelldateien des Add-Ins hostet. Dadurch werden Fehler wegen websiteübergreifendem Skripting vermieden. Ihr Server kann die Anforderung an andere Webdienste als Proxy weiterleiten, wenn Ihre Anwendung dies erfordert.

Das Hinzufügen eines Identitätstokens zur Dienstanforderung, die Ihr Add-In sendet, ist einfach. Sie fordern das Token an, verwenden es und arbeiten dann mit der Antwort des Webdiensts. Das folgende Beispiel zeigt ein einfaches XML-Dokument, das Sie mit der  **XmlHttpRequest** -Methode an Ihren Server senden.

## Anfordern eines Tokens von Ihrem Exchange-Server


Diese einfache Initialisierungsmethode für ein Add-In verwendet die Methode  **getUserIdentityTokenAsync** zum Anfordern eines Identitätstokens vom Exchange-Server. Der Parameter _getUserIdentityToken_ ist die Funktion, die aufgerufen wird, wenn die asynchrone Anforderung an den Server zurückgegeben wird. Informationen zur Rückrufmethode finden Sie im nächsten Schritt.


```
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## Verwenden des Identitätstokens


Die Rückruffunktion der  **getUserIdentityTokenAsync** -Methode hat einen Parameter, der das Benutzeridentitätstoken in seiner **value** -Eigenschaft enthält.

Diese Rückruffunktion erstellt ein  **XMLHttpRequest** -Objekt, um den Webdienst aufzurufen. Legen Sie die Eigenschaft **onreadystatechange** für das **XMLHttpRequest** -Objekt auf den Namen der Funktion fest, die ausgeführt werden soll, wenn Ihr Add-In eine Antwort vom Webdienst erhält.




```
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; CHARSET=Windows-1252");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## Verwenden der Webdienstantwort


Dies ist eine weitere einfache Funktion, die die Antwort des Webdiensts verarbeitet. Sie folgt dem Standardmuster für  **XHMHttpResponse** -Rückruffunktionen. Die Funktion wartet, bis die gesamte Antwort vom Webdienst empfangen wurde, und legt anschließend den Inhalt der Antwort auf der Benutzeroberfläche des Add-ins ab. Die Antwort, die diese Funktion analysiert, ist die Antwort des Webdiensts. Informationen zu dieser Antwort finden Sie unter [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md). 


```
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## Erste Schritte mit dem Aufrufen eines Webdiensts mithilfe von Identitätstoken


Identitätstoken bieten Informationen zur Identität des Clients, der Ihren Dienst für einen Webdienst aufruft, der auf Ihrem Server ausgeführt wird. Zum Verwenden von Identitätstoken benötigen Sie Folgendes:


- Ein Outlook-Add-In, das ein Identitätstoken vom Exchange-Server anfordert und es an Ihren Webdienst sendet. Anhand der Informationen in diesem Thema können Sie dieses Add-In erstellen.
    
- Einen Webdienst, der auf dem Server ausgeführt wird, der die Benutzeroberfläche für Ihr Add-In bereitstellt, auf der das Identitätstoken auf Gültigkeit überprüft wird. In einem der folgenden Themen finden Sie die Informationen, die Sie zum Erstellen des Webdiensts benötigen:
    
      - [Verwenden der Exchange-Tokenüberprüfungsbibliothek](../outlook/use-the-token-validation-library.md) – Wenn Sie die Validierungsbibliothek verwenden, die wir bereitstellen.
    
  - [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md) – Wenn Sie eigenen Validierungscode schreiben.
    

### Code für das Beispiel-Add-In


Die folgenden Dateien sind für das in diesem Artikel behandelte Add-In erforderlich:


- IdentityTest.js: Die JavaScript-Datei, die die Geschäftslogik für das Add-In bereitstellt
    
- IdentityTest.html: Die HTML-Datei, die die Benutzeroberfläche für das Add-In bereitstellt
    
Sie benötigen außerdem den Webdienst zum Identitätstest. Informationen zu diesem Webdienst finden Sie unter [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md).


#### IdentityTest.js

Das folgende Beispiel zeigt die Datei  **IdentityTest.js**.


```
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; CHARSET=Windows-1252");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### IdentityTest.html

Das folgende Beispiel zeigt die Datei  **IdentityTest.html**.


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## Nächste Schritte


Sie wissen jetzt, wie ein Identitätstoken angefordert wird, und müssen nun das Token auf der Serverseite der Anforderung verwenden. Die folgenden Artikel unterstützen Sie bei den ersten Schritten:


- [Verwenden der Exchange-Tokenüberprüfungsbibliothek](../outlook/use-the-token-validation-library.md)
    
- [Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md)
    
- [Authentifizieren eines Benutzers mit einem Identitätstoken für Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## Zusätzliche Ressourcen



- [Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken](../outlook/authentication.md)
    
- [Innenleben des Exchange-Identitätstokens](../outlook/inside-the-identity-token.md)
    
