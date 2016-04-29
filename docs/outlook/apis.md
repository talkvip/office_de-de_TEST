
# Outlook-Add-In-APIs
Informationen zum Angeben der für Ihr Outlook-Add-In erforderlichen APIs.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Um APIs in Ihrem Outlook-Add-In zu verwenden, müssen Sie den Speicherort der Office.js-Bibliothek, den Anforderungssatz, das Schema und die Berechtigungen angeben.

## Office.js

Um mit der Outlook-Add-In-API zu interagieren, müssen Entwickler die JavaScript-APIs in Office.js verwenden. Das CDN für die Bibliothek ist  _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_. Add-Ins, die an den Informationsspeicher übermitteln werden, müssen anhand dieses CDN auf "Office.js" verweisen, sie können keinen lokalen Verweis verwenden. Deklarieren Sie das CDN im **[<head>]** -Tag der Webseite (HTML-, ASPX- oder PHP-Datei), die die Benutzeroberfläche Ihres Add-Ins implementiert, im **[src]** -Attribut des **[script]** -Tags:


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Beim Hinzufügen neuer APIs bleibt die URL zuOffice.js gleich. Die Version in der URL wird nur geändert, wenn eines bestehendes API-Verhalten unterbrochen werden soll.

Die Bibliothek muss innerhalb von 5 Sekunden nach dem Starten des Add-Ins geladen werden; andernfalls bestimmt Outlook, dass die Seite nicht reagiert, und zeigt eine Fehlermeldung an.


## Anforderungssätze

Alle Outlook-APIs gehören zu dem Postfachanforderungssatz. Der Postfachanforderungssatz weist Versionen auf, und jeder neue Satz von APIs, der veröffentlicht wird, gehört zu einer höheren Version des Satzes. Nicht alle Outlook-Clients unterstützen den neuesten Satz von APIs, wenn diese freigegeben werden, aber wenn ein Outlook-Client Unterstützung für einen Anforderungssatz deklariert, unterstützt er alle APIs in diesem Anforderungssatz. 

Durch Angeben einer Mindestversion für den Anforderungsstz im Manifest wird gesteuert, in welchem Outlook-Client das Add-In angezeigt wird. Wenn beispielsweise Version 1.3 für den Anforderungssatz angegeben wird, bedeutet dies, dass das Add-In in keinem Outlook-Client angezeigt wird, der nicht mindestens Version 1.3 unterstützt. Durch Angeben eines Anforderungssatzes wird Ihr Add-In jedoch nicht auf die APIs in dieser Version beschränkt. Wenn das Add-In Version 1.1 des Anforderungssatzes angibt, jedoch in einem Outlook-Client ausgeführt wird, der Version 1.3 unterstützt, kann das Add-In trotzdem diese neuen APIs verwenden. Der Anforderungssatz steuert ledliglich, in welchen Outlook-Clients das Add-In angezeigt wird.

Um die Verfügbarkeit von APIs aus einem Anforderungssatz zu überprüfen, der höher als der im Manifest angegebene ist, können Sie die standardmäßige JavaScript-Technik verwenden:




```
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction…  
}

```

Für APIs, die sich in der im Manifest angegebenen Version des Anforderungssatzes befinden, sind solche Überprüfungen nicht erforderlich.

Entwickler sollten den minimalen Anforderungssatz angeben, der den kritischen Satz von APIs für ihr Szenario unterstützt, ohne den die entscheidenen Features des Add-Ins nicht funktionieren. Der Anforderungssatz im Manifest wird in den Elementen  **Requirements**, **Sets** und **Set** angegeben. Weitere Informationen finden Sie unter [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md).

Das Element  **Methods** gilt nicht für Mail-Add-Ins. Deshalb können Sie keine Unterstützung für bestimme Methoden deklarieren.


## Berechtigungen

Ihr Add-In benötigt die entsprechenden Berechtigungen, damit die erforderlichen APIs verwendet werden können. Es gibt vier Stufen von Berechtigungen, die nachfolgende zusammengefasst werden. Einzelheiten finden Sie unter [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../../docs/outlook/privacy/understanding-outlook-add-in-permissions.md).


|
|
|**Berechtigungsstufe**|**Beschreibung**|
|:-----|:-----|
|Eingeschränkt|Ermöglicht die Verwendung von Entitäten, jedoch nicht von regulären Ausdrücken.|
|Element lesen|Zusätzlich zu den in  _Eingeschränkt_ zulässigen Aktionen ist Folgendes erlaubt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Reguläre Ausdrücke</p></li><li><p>Lesezugriff für Outlook-Add-In-API</p></li><li><p>Abrufen der Elementeigenschaften und des Rückruftokens</p></li></ul>|
|Lese-/Schreibzugriff|Zusätzlich zu den in  _Restricted_ zulässigen Aktionen ist Folgendes zulässig:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Vollständiger Zugriff auf Outlook-Add-In-API, außer <span class="keyword">makeEwsRequestAsync</span></p></li><li><p> Festlegen der Elementeigenschaften</p></li></ul>|
|Postfach lesen/schreiben|Zusätzlich zu den in  _Lese-/Schreibzugriff_ zulässigen Aktionen ist Folgendes zulässig:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Erstellen, Lesen, Schreiben von Elementen und Ordnern</p></li><li><p>Senden von Elementen</p></li><li><p>Aufrufen von <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html.aspx#makeEwsRequestAsync" target="_blank">makeEwsRequestAsync</a>,</p></li></ul>|
Im Allgemeinen sollten Sie die für Ihr Add-In erforderliche Mindestberechtigungen angeben. Berechtiungen werden im  **Permissions** -Element im Manifest deklariert. Weitere Informationen finden Sie unter [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md). Informationen zu Sicherheitsproblemen finden Sie unter [Datenschutz, Berechtigungen und Sicherheit für Outlook-Add-Ins](../../docs/develop/privacy-and-security.md)


## Zusätzliche Ressourcen



- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Outlook Add-in API](http://dev.outlook.com/reference/add-ins/index.mdl.aspx)
    
- [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md)
    
- [Datenschutz, Berechtigungen und Sicherheit für Outlook-Add-Ins](../../docs/develop/privacy-and-security.md)
    
