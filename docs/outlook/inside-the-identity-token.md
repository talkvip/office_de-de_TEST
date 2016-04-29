
# Innenleben des Exchange-Identitätstokens
In diesem Thema erfahren Sie, was sich innerhalb eines Exchange 2013-Identitätstokens befindet.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Das Authentifizierungsidentitätstoken, das der Exchange-Server an Ihr Outlook-Add-in sendet, ist für Ihr Add-in nicht transparent. Sie müssen nicht in das Token hineinsehen, um es an Ihren Server weiterzusenden. Wenn Sie allerdings den Webdienstcode schreiben, der mit Ihrem Outlook-Add-in interagiert, müssen Sie den Inhalt des Identitätstokens kennen.

## Was ist ein Identitätstoken?


Ein Identitätstoken ist eine Base64-URL-codierte Zeichenfolge, die vom Exchange-Server, der das Identitätstoken gesendet hat, selbstsigniert wird. Das Token ist nicht verschlüsselt, und der öffentliche Schlüssel, den Sie zum Überprüfen der Signatur verwenden, wird auf dem Exchange-Server gespeichert, der das Token ausgestellt hat. Das Token besteht aus drei Bestandteilen: einem Header, einer Ladung und einer Signatur. In der Tokenzeichenfolge werden die Bestandteile durch das Zeichen "." voneinander getrennt, damit Sie das Token problemlos aufteilen können.

Exchange 2013 verwendet ein JSON Web Token (JWT) für das Identitätstoken. Informationen zu JWT-Token finden Sie im [JSON Web Token (JWT)-Internetentwurf](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl).


### Identitätstokenheader

Der Header identifiziert das Token und teilt dem Webdienst mit, welche Art von Token präsentiert wird. Das folgende Beispiel veranschaulicht den Header des Tokens.

{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }In der folgenden Tabelle werden die Bestandteile des Identitätstokenheaders beschrieben.


**Bestandteile des Identitätstokenheaders**


|**Anspruch**|**Wert**|**Beschreibung**|
|:-----|:-----|:-----|
|typ|"JWT"|Identifiziert das Token als JSON Web Token (JWT). Alle vom Exchange-Server bereitgestellten Identitätstoken sind JWT-Token.|
|alg|"RS256"|Der zum Erstellen der Signatur verwendete Hashalgorithmus. Alle vom Exchange-Server bereitgestellten Token verwenden den RS-256-Algorithmus.|
|x5t|Zertifikatfingerabdruck|Der X.509-Fingerabdruck des Tokens.|

### Identitätstokenlast

Die Last enthält die Authentifizierungsansprüche, mit denen das E-Mail-Konto und der Exchange-Server, der das Token gesendet hat, identifiziert werden. Das folgende Beispiel veranschaulicht den Abschnitt für die Last.

{ "aud" : "https://mailhost.contoso.com/IdentityTest.html", "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", "nbf" : "1331579055", "exp" : "1331607855", "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com", "isbrowserhostedapp":"true", "appctx" : { "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" } }In der folgenden Tabelle werden die Bestandteile der Identitätstokenlast beschrieben.


**Bestandteile der Identitätstokenlast**


|**Anspruch**|**Beschreibung**|
|:-----|:-----|
|aud|Die URL des Add-Ins, welches das Token angefordert hat. Ein Token ist nur gültig, wenn es über das Add-In gesendet wird, das im Browser des Clients ausgeführt wird. Wenn das Add-In das Office-Add-Ins **REMOVE_ME** -Manifestschema v1.1 verwendet, ist diese URL die URL, die im ersten **SourceLocation**-Element unter dem Formattyp  **ItemRead** oder **ItemEdit** angegeben wird, je nachdem, welches zuerst als Teil des Elements [FormSettings](http://msdn.microsoft.com/de-de/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) im Add-In-Manifest auftritt.|
|iss|Ein eindeutiger Bezeichner für den Exchange-Server, der das Token ausgestellt hat. Alle von diesem Exchange-Server ausgestellten Token weisen denselben Bezeichner auf.|
|nbf|Das Datum und die Uhrzeit, ab dem bzw. der das Token gültig ist. Dieser Wert bezeichnet die Anzahl der Sekunden seit dem 1. Januar 1970. |
|exp|Das Datum und die Uhrzeit, bis zu dem bzw. der das Token gültig ist. Dieser Wert bezeichnet die Anzahl der Sekunden seit dem 1. Januar 1970.|
|appctxsender|Ein eindeutiger Bezeichner für den Exchange-Server, der den Anwendungskontext gesendet hat.|
|isbrowserhostedapp|Gibt an, ob das Add-in in einem Browser gehostet wird.|
|appctx|Der Anwendungskontext für das Token. |
Die Informationen im "appctx" liefern die Adresse des E-Mail-Kontos und einen eindeutigen Bezeichner für das Konto. In der folgenden Tabelle werden die Bestandteile des "appctx"-Anspruchs beschrieben.



|**Bestandteil des "appctx"-Anspruchs**|**Beschreibung**|
|:-----|:-----|
|msexchuid|Ein eindeutiger Bezeichner für das E-Mail-Konto und den Exchange-Server.|
|version|Die Versionsnummer des Tokens. Für alle Token, die von einem Server mit Exchange 2013 bereitgestellt werden, lautet dieser Wert "ExIdTok.V1".|
|amurl|Die URL des Dokuments mit den Authentifizierungsmetadaten, das den öffentlichen Schlüssel des zum Signieren des Tokens verwendeten X.509-Zertifikats enthält. Weitere Informationen zum Verwenden des Dokuments mit den Authentifizierungsmetadaten finden Sie unter [Überprüfen eines Exchange-Identitätstokens](8503a3e8-458a-4a4e-9e95-65cd7bb1954d.md).|

### Identitätstokensignatur

Die Signatur wird erstellt durch das Hashing der Header- und Lastabschnitte mit dem im Header angegebenen Algorithmus und mithilfe des selbstsignierten X509-Zertifikats, das sich auf dem Server in dem für die Last angegebenen Speicherort befindet. Ihr Webdienst kann diese Signatur überprüfen, um sicherzustellen, dass das Identitätstoken von dem Server stammt, an den es gesendet werden soll.


## Weitere Ressourcen



- [Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken](c0520a1e-d9ba-495a-a99f-6816d7d2a23e.md)
    
- [Aufrufen eines Diensts von einem Outlook-Add-In mithilfe eines Identitätstokens in Exchange](15c1a27e-6a75-411b-bf56-046d0eba3b48.md)
    
- [Verwenden der Exchange-Tokenüberprüfungsbibliothek](f7f4813a-3b2d-47bb-bf93-71b64620a56b.md)
    
- [Überprüfen eines Exchange-Identitätstokens](8503a3e8-458a-4a4e-9e95-65cd7bb1954d.md)
    
