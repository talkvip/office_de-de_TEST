
# Voraussetzungen zum Ausführen von Office-Add-ins
Erfahren Sie mehr über die Software- und Geräteanforderungen zum Ausführen von Office-Add-Ins.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_


## Serveranforderungen

Zum Installieren und Ausführen eines Office-Add-Ins müssen Sie zuerst die Manifest- und Webseitendatei für die Benutzeroberfläche und den Code Ihres Add-Ins an den entsprechenden Speicherorten auf dem Server bereitstellen.

Bei allen Add-In-Typen (Inhalt, Outlook und Aufgabenbereich) müssen Sie die Webseitendatei Ihres Add-ins auf einem Webserver oder bei einem Webhostingdienst wie [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) bereitstellen.


 >**Hinweis**   Beim Entwickeln und Debuggen eines Add-Ins in Visual Studio stellt Visual Studio die Webseitendatei Ihres Add-Ins lokal über IIS Express bereit und führt sie aus. Daher wird kein zusätzlicher Webserver benötigt. Gleichermaßen wird die Webseitendatei beim Entwickeln und Debuggen mit Napa im Browser im Speicher bereitgestellt, der mit Ihrem Konto zur Anmeldung bei Napa verknüpft ist, und von dort aus ausgeführt.

Für Inhalts- und Aufgabenbereich-Add-ins in den unterstützten Office-Hostanwendungen (Access Web Apps, Word, Excel, PowerPoint und Project) benötigen Sie zum Hochladen der XML-Manifest-Datei des Add-ins außerdem entweder eine [Netzwerkdateifreigabe](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) oder einen [Add-in-Katalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) auf SharePoint.

Zum Testen und Ausführen eines Outlook-Add-Ins muss sich das Outlook-E-Mail-Konto des Benutzers auf Exchange 2013 oder höher befinden, das über Office 365, Exchange Online oder als lokale Installation verfügbar ist. Der Benutzer oder Administrator installiert Manifestdateien für Outlook-Add-Ins auf diesem Server.

Office-Add-Ins werden von POP- und IMAP-E-Mail-Konten in Outlook nicht unterstützt.


## Zusammenfassung zur Clientunterstützung

In der folgenden Tabelle sind die Office-Hostanwendungen (einschließlich Desktop-, Tablet- und Webclients), die Office-Add-Ins ausführen können, und die von jedem Host unterstützten Add-In-Typen aufgeführt.


**Unterstützte Add-in-Typen**


|**Office-Anwendung**|**Inhalts-Add-ins**|**Outlook-Add-Ins**|**Aufgabenbereich-Add-ins**|
|:-----|:-----|:-----|:-----|
|Access-Web-Apps|
![Häkchen](../../images/mod_off15_checkmark.png)

|||
|Excel 2013 oder höher|
![Häkchen](../../images/mod_off15_checkmark.png)

||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Excel Online|
![Häkchen](../../images/mod_off15_checkmark.png)

||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Excel für iPad|
![Häkchen](../../images/mod_off15_checkmark.png)

||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Outlook 2013 oder höher||
![Häkchen](../../images/mod_off15_checkmark.png)

||
|Outlook für Mac||
![Häkchen](../../images/mod_off15_checkmark.png)

||
|Outlook Web App||
![Häkchen](../../images/mod_off15_checkmark.png)

||
|OWA für mobile Geräte||
![Häkchen](../../images/mod_off15_checkmark.png)

||
|PowerPoint 2013 oder höher|
![Häkchen](../../images/mod_off15_checkmark.png)

||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|PowerPoint Online|
![Häkchen](../../images/mod_off15_checkmark.png)

||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Project 2013 oder höher|||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Word 2013 oder höher|||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Word Online|||
![Häkchen](../../images/mod_off15_checkmark.png)

|
|Word für iPad|||
![Häkchen](../../images/mod_off15_checkmark.png)

|

## Clientanforderungen: Windows-Desktop und Tablet

Die folgende Software ist zum Entwickeln eines Office-Add-Ins für die unterstützten Office-Desktopclients oder Webclients erforderlich, die auf Desktop-, Laptop- oder Tablet-Geräten unter Windows ausgeführt werden:


- Für x86- und x64-Windows-Desktops sowie für Tablets wie das Surface Pro:
    
      - Die 32- oder 64-Bit-Version von Office 2013 oder höher unter Windows 7 oder höher
    
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013 oder eine höhere Version des Office-Clients, falls Sie ein Office-Add-In speziell für einen dieser Office-Desktopclients testen und ausführen. Office-Desktopclients können lokal oder über Klick-und-Los auf dem Clientcomputer installiert werden.
    
- Internet Explorer 9 oder höher muss installiert, aber nicht der Standardbrowser sein. Zur Unterstützung von Office-Add-Ins verwendet der Office-Client, der als Host agiert, Browserkomponenten, die Bestandteil von Internet Explorer 9 oder höher sind.
    
- Eine der folgenden Anwendungen als Standardbrowser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, oder eine höhere Version eines dieser Browser.
    
- Ein HTML- und JavaScript-Editor wie z. B. Editor, [Visual Studio und die Microsoft Developer Tools ](https://www.visualstudio.com/features/office-tools-vs) oder ein Webentwicklungstool eines anderen Anbieters.
    

## Clientanforderungen: OS X-Desktop

Outlook für Mac wird als Bestandteil von Office 365 verteilt und unterstützt Outlook-Add-Ins. Für die Ausführung von Outlook-Add-Ins in Outlook für Mac gelten die gleichen Anforderungen wie für Outlook für Mac selbst: Als Betriebssystem muss OS X v10.10 "Yosemite" oder höher verwendet werden. Da Outlook für Mac WebKit als Layoutmodul zum Rendern der Add-In-Seiten verwendet, besteht keine zusätzliche Browserabhängigkeit.


## Clientanforderungen: Browserunterstützung für Office Online-Webclients und SharePoint

Jeder Browser, der ECMAScript 5.1, HTML5 und CSS3 unterstützt, z. B. Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 oder eine spätere Version dieser Browser.


## Clientanforderung: Smartphone und Tablet (nicht unter Windows)

Wenn OWA für mobile Geräte und Outlook Web App in einem Browser auf einem Smartphone oder einem anderen als einem Windows Tablet-Gerät ausgeführt wird, ist die folgende Software zum Testen und Ausführen von Outlook-Add-Ins erforderlich.



|**Hostanwendung**|**Gerät**|**Betriebssystem**|**Exchange-Konto**|**Mobiler Browser**|
|:-----|:-----|:-----|:-----|:-----|
|OWA für Android|Android Smartphones. Im technischen Sinne gelten diese Geräte für das [Android-BS](http://developer.android.com/guide/practices/screens_support.mdl) als "klein" oder "normal".|Android 4.4 KitKat oder höher|Beim letzten Update von Office 365 für Unternehmen oder Exchange Online|Natives Add-in für Android, Browser nicht anwendbar|
|OWA für iPad|iPad 2 oder höher|iOS 5 oder höher|Beim letzten Update von Office 365 für Unternehmen oder Exchange Online|Natives Add-in für iOS, Browser nicht anwendbar|
|OWA für iPhone|iPhone 4S oder höher|iOS 6 oder höher|Beim letzten Update von Office 365 für Unternehmen oder Exchange Online|Natives Add-in für iOS, Browser nicht anwendbar|
|Outlook Web App|iPhone 4 oder höher, iPad 2 oder höher, iPod Touch 4 oder höher|iOS 5 oder höher|Bei Office 365, Exchange Online oder lokal auf Exchange Server 2013 oder höher|Safari|

## Komponenten einer Office-Add-In-Lösung


Eine typische Office-Add-In-Lösung enthält die folgenden Komponenten:


- Ein Clientgerät, auf dem der unterstützte Office-Client ausgeführt wird. Es kann sich hierbei um einen Desktop, einen Laptop, ein Tablet oder ein Smartphone handeln (für Outlook-Add-Ins auf OWA für mobile Geräte). 
    
- Für Access Web Apps, Word, Excel, PowerPoint oder Project:
    
      - Eine Datenbank, ein Dokument, eine Arbeitsmappe, eine Präsentation oder ein Projekt.
    
  - Ein Aufgabenbereich- oder Inhalts-Add-in, das der Benutzer über den öffentlichen Office Store oder über einen privaten SharePoint- oder dateibasierten Add-in-Katalog installiert hat.
    
- Für Outlook: 
    
      - Das E-Mail-Konto und das Postfach des Benutzers, die sich auf einem Exchange Server befinden.
    
  - Ein Outlook-Add-In, das der Benutzer oder Exchange Server-Administrator über das Exchange Admin Center (EAC) installiert hat.
    

 >**Hinweis**  Die Installation eines Office-Add-Ins durch den Benutzer besteht aus einem Zeiger auf die entsprechende XML-Manifestdatei, mit dem die URL zum Laden der Add-In-Webseite und des Skripts zur Laufzeit angegeben wird.

Für alle unterstützten Office-Anwendungen besteht die Implementierung des Office-Add-Ins selbst aus den folgenden serverbasierten Komponenten:


- Eine XML-Manifestdatei, die in einem öffentlichen oder privaten Add-in-Katalog oder im Exchange Server des Benutzers gespeichert ist.
    
- Die HTML-, CSS- und JavaScript-Dateien des Add-ins, die der Benutzer erstellt und die auf einem Webserver gespeichert sind.
    
- Die JavaScript-Bibliotheksdateien, z. B. die JavaScript-API für Office (Office.js) und die Microsoft AJAX-Bibliothek (MicrosoftAjax.js), die von Microsoft bereitgestellt werden. Das Add-in greift über die in der HTML-Datei angegebenen CDN-URLs (Content Delivery Network, Netzwerk für die Inhaltsübermittlung) auf die JavaScript-Bibliotheksdateien zu.
    
Wenn eine unterstützte Office-Anwendung gestartet wird, werden die XML-Manifeste für die Add-Ins gelesen, die für den Benutzer oder von dem Benutzer installiert wurden. Wenn ein Benutzer anschließend ein Office-Add-In in der Office-Anwendung startet, passiert Folgendes: 


1. Für Access Web Apps, Word, Excel, PowerPoint oder Project: Wenn ein Benutzer das Office-Add-In einfügt oder eine Access-Web-App, ein Dokument, eine Arbeitsmappe, eine Präsentation oder ein Projekt öffnet, das bzw. die bereits ein Add-In enthält, wird das Add-In von der Office-Anwendung geladen, und die Benutzeroberfläche des Add-Ins wird angezeigt
    
    Für Outlook: Wenn der aktuelle Outlook-Kontext die Aktivierungsbedingungen eines Add-Ins erfüllt, aktiviert Outlook das Add-In und zeigt das Add-In auf der Outlook-Benutzeroberfläche für eine Auswahl an.
    
2. Für Windows- oder Web-basierte Office-Anwendungen: Die Office-Anwendung öffnet die HTML-Seite in einem Webbrowser-Steuerelement (Desktopclient oder ARM-spezifischer Client) oder einem  **iframe** (Webclient). Das Webbrowser-Steuerelement verwendet Internet Explorer 9-Komponenten (oder höher) und bietet Sicherheit und Leistungsisolation.
    
    Für Outlook für Mac auf OS X-Basis: Outlook für Mac verwendet einen Sandkasten-Hostprozess für WebKit-Laufzeit zum Öffnen der HTML-Seite eines Outlook-Add-Ins, um ähnliche Sicherheits- und Leistungsstufen bereitzustellen.
    
3. Das entsprechende Browsersteuerelement,  **iframe**, oder der Hostprozess für Webkit-Laufzeit lädt den HTML-Textkörper und ruft den Ereignishandler für das  **onload**-Ereignis auf.
    
4. Das Office-Add-Ins-Framework ruft den Ereignishandler für das [initialize](http://msdn.microsoft.com/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignis des [Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx)-Objekts auf.
    
5. Wenn der HTML-Textkörper vollständig geladen wurde und das Add-in initialisiert wurde, kann die Hauptfunktion des Add-ins fortgesetzt werden.
    

## Zusätzliche Ressourcen



- [Office-Add-Ins-Plattformübersicht](../../docs/develop/privacy-and-security.md)
    
