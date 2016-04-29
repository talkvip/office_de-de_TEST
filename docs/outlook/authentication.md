
# Authentifizieren eines Outlook-Add-Ins mithilfe von Exchange-Identitätstoken
Hier erfahren Sie, wie Sie mithilfe von Exchange 2013-Identitätstoken Ihr Outlook-Add-In authentifizieren.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Ihr Outlook-Add-In in kann Ihren Kunden Informationen von überall im Internet liefern, sei es vom Server, auf dem das Add-in gehostet ist, von Ihrem internen Netzwerk oder von einer anderen Ressource in der Cloud. Falls diese Informationen jedoch geschützt sind, muss Ihr Mail-Add-in das Exchange-E-Mail-Konto auf irgendeine Weise Ihrem Informationsdienst zuordnen können. Exchange 2013 kann einmaliges Anmelden (Single Sign-On, SSO) für Ihr Add-in aktivieren, indem ein Token zur Verfügung gestellt wird, welches das E-Mail-Konto identifiziert, von dem aus die Anforderung erfolgt. Sie können dieses Token einem für Ihre Anwendung registrierten Benutzer zuordnen, damit dieser erkannt wird, wenn das Add-In eine Verbindung mit Ihrem Dienst herstellt.

## Identitätstoken


Zwei unserer Beispiel-Add-ins verwenden öffentlich verfügbare Informationen. Mit einem Add-in wird eine Bing-Landkarte für Adressen in einer Nachricht angezeigt, und mit dem anderen Add-in wird eine Vorschau für Links zu YouTube-Videos in einer Nachricht angezeigt. Ihr Mail-Add-in kann jedoch auch auf nicht öffentliche Informationen zugreifen. Sie können den Server, auf dem Ihr Add-In gehostet wird, verwenden, um Ihr Add-in mit den Informationen in Ihrem internen Netzwerk oder in anderen Ressourcen in der Cloud zu verknüpfen.

Sie können viele verschiedene Verfahren zum Identifizieren und Authentifizieren der Add-in-Benutzer verwenden. Exchange 2013 vereinfacht die Benutzerauthentifizierung, indem Ihrem Add-In ein Identitätstoken zur Verfügung gestellt wird, das ein bestimmtes Exchange-E-Mail-Konto identifiziert. In Ihrem Add-in können Sie dieses Token einem registrierten Benutzer zuordnen und somit SSO für Ihre Kunden aktivieren, die Outlook-Add-Ins verwenden. 

Zur Verwendung von SSO in Ihrem Add-in werden vom Code die folgenden Aktionen ausgeführt:


- Eine Funktion in der Outlook-Add-In-API wird aufgerufen, mit der ein Identitätstoken zurückgegeben wird.
    
- Das Token wird zusammen mit einer Anforderung an Ihren Server gesendet.
    
- Die Antwort vom Server wird entpackt, um Informationen von Ihrem Dienst anzuzeigen.
    
Serverseitig ist dieser Vorgang etwas komplexer. Wenn der Server eine Anforderung von einem Outlook-Add-In-Add erhält, passiert Folgendes:


- Das Token wird vom Server überprüft. Sie können unsere [verwaltete Tokenüberprüfungsbibliothek](http://msdn.microsoft.com/de-de/library/f7f4813a-3b2d-47bb-bf93-71b64620a56b%28Office.15%29.aspx) verwenden oder [Ihre eigene Bibliothek](http://msdn.microsoft.com/de-de/library/8503a3e8-458a-4a4e-9e95-65cd7bb1954d%28Office.15%29.aspx) für Ihren Dienst erstellen.
    
- Der Server sucht den eindeutigen Bezeichner aus dem Token, um festzustellen, ob dieser einer bekannten Identität zugeordnet ist. Ihr Dienst muss [eine Methode zum Abgleichen des Bezeichners](http://msdn.microsoft.com/de-de/library/bb28ca39-1780-4162-a899-7be5825beb8e%28Office.15%29.aspx) mit bekannten Benutzern Ihres Diensts implementieren.
    
- Stimmt der eindeutige Bezeichner mit einem Bezeichner überein, der bereits mit einem Satz von Anmeldeinformationen auf dem Server gespeichert wurde, kann Ihr Server die angeforderten Informationen zurückgeben, ohne dass Ihr Kunde sich bei Ihrem Dienst anmelden muss.
    
- Falls der eindeutige Bezeichner unbekannt ist, sendet der Server eine Antwort, in der der Benutzer aufgefordert wird, sich mit Anmeldeinformationen für den Server anzumelden.
    
- Stimmen die Anmeldeinformationen mit einer bekannten Identität auf dem Server überein, können Sie diese Identität dem eindeutigen Bezeichner im Token zuordnen, damit der Server bei der nächsten eingehenden Anforderung antworten kann, ohne dass ein zusätzlicher Schritt für die Anmeldung erforderlich ist.
    

(../../images  Dies ist lediglich ein möglicher Vorschlag für die Verwendung des Identitätstokens. Wie immer müssen Sie im Zusammenhang mit Identität und Authentifizierung sicherstellen, dass Ihr Code die Sicherheitsanforderungen Ihres Unternehmens erfüllt.

Lassen Sie uns das einmal genauer ansehen. Als Beispiel verwenden wir ein einfaches Outlook-Add-In, mit dem das Identitätstoken und eine Liste von Telefonnummern aus der Nachricht an einen Webdienst gesendet werden. 


## Inhalt dieses Abschnitts




|**Artikel**|**Beschreibung**|
|:-----|:-----|
|[Innenleben des Exchange-Identitätstokens](../outlook/inside-the-identity-token.md)|Beschreibt die speziellen Ansprüche, die im Token enthalten sind.|
|[Aufrufen eines Diensts von einem Outlook-Add-In mithilfe eines Identitätstokens in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)|Enthält Codebeispiele für Outlook-Add-In-Autoren.|
|[Verwenden der Exchange-Tokenüberprüfungsbibliothek](../outlook/use-the-token-validation-library.md)|Enthält Codebeispiele zum Schreiben von serverseitigem Code mithilfe der .NET Framework-Validierungsbibliothek.|
|[Überprüfen eines Exchange-Identitätstokens](../outlook/validate-an-identity-token.md)|Enthält Codebeispiele zum Implementieren Ihrer eigenen Tokenüberprüfung.|
|[Authentifizieren eines Benutzers mit einem Identitätstoken für Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)|Enthält Codebeispiele, um ein einfaches System für einmaliges Anmelden in einen Dienst zu implementieren.|

## Zusätzliche Ressourcen



- [Outlook-Add-Ins](../outlook/outlook-add-ins.md)
    
- [Aufrufen von Webdiensten aus einem Outlook-Add-In](../outlook/web-services.md)
    


