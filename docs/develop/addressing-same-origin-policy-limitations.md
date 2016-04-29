
# Behandeln von Richtlinieneinschränkungen aufgrund desselben Ursprungs in Office-Add-ins
Dieser Artikel bietet eine kurze Übersicht über Möglichkeiten zur Umgehung der Einschränkungen durch die "Richtlinie des gleichen Ursprungs", die verhindert, dass ein von einer Domäne geladenes Skript Eigenschaften eines Dokuments von einer anderen Domäne in einem Office-Add-In abruft oder verändert.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Die vom Browser durchgesetzte "Richtlinie des gleichen Ursprungs" verhindert, dass ein von einer Domäne geladenes Skript Eigenschaften einer Webseite aus einer anderen Domäne abruft oder verändert. Das heißt, dass die Domäne einer angeforderten URL standardmäßig identisch mit der Domäne der aktuellen Webseite sein muss. Beispielsweise verhindert diese Richtlinie, dass eine Webseite in einer anderen Domäne als der, in der sie gehostet wird, [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/)-Webdienstauffrufe ausführt.

Da Office-Add-Ins in einem Browsersteuerelement gehostet werden, gilt die „Richtlinie des gleichen Ursprungs" auch für Skripts, die auf ihren Webseiten ausgeführt werden.

Es gibt mehrere Möglichkeiten, bei der Entwicklung von Add-ins die Durchsetzung der „Richtlinie des gleichen Ursprungs" zu umgehen:

- Verwenden von JSON/P für den anonymen Zugriff 
    
- Implementieren von serverseitigen Skripts unter Verwendung eines tokenbasierten Authentifizierungsschemas
    
- Verwenden von Cross-Origin Resource Sharing (CORS)
    
- Erstellen Ihres eigenen Proxys unter Verwendung von IFRAME und POST MESSAGE
    

## Verwenden von JSON/P für den anonymen Zugriff


Eine Möglichkeit zur Umgehung dieser Einschränkung ist die Verwendung von JSON/P, um einen Proxy für den Webdienst bereitzustellen. Zu diesem Zweck fügen Sie ein  **script**-Tag mit einem  **src**-Attribut ein, das auf Skript verweist, das in einer beliebigen Domäne gehostet wird. Die  **script**-Tags können programmgesteuert erstellt werden, und die URL, auf die das  **src**-Attribut verweist, wird dynamisch erstellt. Anschließend werden über URI-Abfrageparameter Parameter an die URL übergeben. Webdienstanbieter erstellen und hosten JavaScript-Code unter bestimmten URLs und geben in Abhängigkeit von den URI-Abfrageparametern unterschiedliche Skripts zurück. Diese Skripts werden dann dort ausgeführt, wo sie eingefügt werden und funktionieren wie erwartet.

Im Folgenden ist ein Beispiel für JSON/P gezeigt, in dem ein Verfahren verwendet wird, das in jedem beliebigen Office-Add-In funktioniert.




```
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## Implementieren von serverseitigen Skripts unter Verwendung eines tokenbasierten Authentifizierungsschemas


Eine weitere Möglichkeit zur Umgehung der Einschränkungen durch die „Richtlinie des gleichen Ursprungs" ist die Implementierung der Webseite des Add-ins als ASP-Seite, die OAuth verwendet oder Anmeldeinformationen in Cookies zwischenspeichert.

Ein Beispiel zur Verwendung von OAuth zu Authentifizierung finden Sie unter [Einen SharePoint-Webpart mit OAuth twittern](http://aidangarnish.net/post/Twitter-SharePoint-Web-Part-With-OAuth.aspx).

Ein Beispiel für serverseitigen Code, der die Verwendung des  **Cookie**-Objekts in  **System.Net** zum Abrufen und Festlegen der Cookiewerte erläutert, finden Sie unter der [Value](http://msdn2.microsoft.com/DE-DE/library/4f772twc)-Eigenschaft.


## Verwenden von Cross-Origin Resource Sharing (CORS)


Ein Beispiel zu Verwendung des Features des Cross-Origin Resource Sharing von [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.mdl) finden Sie im Abschnitt "Cross Origin Resource Sharing (CORS)" auf [Neue Tricks für XMLHttpRequest2](http://www.mdl5rocks.com/de/tutorials/file/xhr2/).


## Erstellen Ihres eigenen Proxys unter Verwendung von IFRAME und POST MESSAGE


Ein Beispiel zur Erstellung Ihres eigenen Proxys unter Verwendung von IFRAME und POST MESSAGE finden Sie unter [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).


## Zusätzliche Ressourcen



- [Datenschutz und Sicherheit bei Office-Add-Ins](87c59a88-10e2-4c88-b6a8-736bd356e5f8.md)
    
