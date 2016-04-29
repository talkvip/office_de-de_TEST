
# Verweisen auf die Bibliothek der JavaScript-API für Office über das Netzwerk für die Inhaltsübermittlung (Content Delivery Network, CDN)
Verweisen Sie auf die Bibliothek der JavaScript-API für Office ( **Office.js** und zugehörige anwendungsspezifische JS-Dateien) von ihrem Speicherort für das Netzwerk für die Inhaltsübermittlung (Content Delivery Network, CDN) aus.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Die Bibliothek der [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx) besteht aus der Datei Office.js und den zugehörigen hostanwendungsspezifischen JS-Dateien, wie beispielsweise Excel-15.js und Outlook-15.js. Beim Entwickeln einer Office-Add-In für eine beliebige Office-Hostanwendung sollten Sie auf die Bibliothek der JavaScript-API für Office innerhalb des `<head>`-Tags der Webseite (z. B. eine HTML-, ASPX- oder PHP-Datei) verweisen, die die Benutzeroberfläche des Add-ins implementiert. Fügen Sie hierfür ein  `script`-Tag hinzu, dessen  `src`-Attribut auf die folgende CDN-URL festgelegt ist:



```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Die  `/1/` vor `office.js` in der CDN-URL gibt an, dass die letzte inkrementelle Version innerhalb von Version 1 von Office.js verwendet wird. Da die JavaScript-API für Office die Abwärtskompatibilität bewahrt, werden in der neuesten Version API-Elemente, die früher in Version 1 eingeführt wurden, weiter unterstützt. Wenn Sie ein vorhandenes Projekt aktualisieren müssen, finden Sie Informationen unter [Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien](641dc473-0931-4e00-8164-e7808ceed64d.md).
Wenn die Add-In zum ersten Mal geladen wird, werden die Bibliotheksdateien der JavaScript-API für Office heruntergeladen und zwischengespeichert, um sicherzustellen, dass die Add-In die aktuelle Implementierung der Datei  **Office.js** und der anwendungsspezifischen JS-Dateien verwendet.
Die standardmäßige Datei "Home.html" in Ihrem Projekt enthält das entsprechende  `script`-Tag, wenn Sie Ihre Add-In mit den  **Add-in für Office**-Projektvorlagendateien erstellen, die mit der neuesten Version von Visual Studio mit dem [aktuellen Update für die Microsoft Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) bereitgestellt werden:

## Zusätzliche Ressourcen



- [Grundlegendes zur JavaScript-API für Office](01180dae-ca45-40c8-b3dd-fd2a85651c0c.md)
    
- [Office-Add-Ins-Plattformübersicht](e64de870-ce22-4331-92e7-76d35279bf91.md)
    
- [Entwicklungslebenszyklus von Office-Add-Ins](c35b4b2b-7869-4501-9f10-888c8e74c98c.md)
    
- [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)
    
