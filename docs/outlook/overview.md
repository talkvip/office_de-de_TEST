
# Übersicht über Architektur und Features von Outlook-Add-Ins
Informationen zu Architektur und Features von Outlook-Add-Ins

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Architektur

Ein Outlook-Add-In besteht aus einem XML-Manifest und Code (JavaScript und HTML). Das Manifest gibt den Namen und eine Beschreibung für das Add-In sowie dessen Integration in Outlook an. Mit dem Manifest können Entwickler die Platzierung von Schaltflächen in Befehlsoberflächen, Links für Übereinstimmungen mit regulären Ausdrücken und anderes deklarieren. Das Manifest definiert auch die URL, die den JavaScript- und HTML-Code für das Add-In hostet.

Wenn ein Benutzer oder Administrator ein Add-In bezieht, wird das Add-In-Manifest im Postfach des Benutzers oder in der Organisation gespeichert. Beim Start von Outlook werden alle Manifeste, die der Benutzer installiert hat, geladen, verarbeitet und die Erweiterungspunkte für das Add-In eingerichtet (z. B. Schaltflächen auf Befehlsoberflächen angezeigt, ein regulärer Ausdruck für die aktuell ausgewählte Nachricht ausgeführt usw.). Jetzt kann der Benutzer das Add-In verwenden.

Wenn der Benutzer mit dem Add-In interagiert, werden die JavaScript- und HTML-Dateien aus dem im Manifest angegebenen Hostspeicherort geladen.

Add-Ins verwenden die API Office.js zum Zugriff auf die Outlook-Add-In-API und interagieren mit Outlook.


**Interaktion typischer Komponenten, wenn der Benutzer Outlook startet**

![Ereignisablauf beim Starten einer Outlook-Mail-App](../../images/olowawecon15_LoadingDOMAgaveRuntime.png)
### Versionsverwaltung

Im Zuge der Weiterentwicklung von Outlook-Clients und der Add-In-Plattform sowie der Entwicklung neuer Möglichkeiten der Integration von Add-Ins können wir manchmal ein Feature nicht gleichzeitig für alle Clients (Mac, Windows, Web, mobile Geräte) implementieren. Aus diesem Grund versehen wir das Manifest und die APIs mit Versionsnummern. So unterstützt die Plattform jederzeit die Abwärtskompatibilität, sodass Entwickler ein Add-In erstellen können, das abwärtskompatibel auf älteren Clients funktioniert, aber gleichzeitig die neuen Features in neueren Clients zur Verfügung stellt. Weitere Informationen zur Funktionsweise der Versionsverwaltung finden Sie in [Outlook-Add-In-Manifeste](98ee2cca-f866-4377-bd1f-6a5338ebd7b1.md).


## Features von Outlook-Add-Ins

Outlook-Add-Ins bieten viele umfangreiche Features, die verwendet werden können, um verschiedene Szenarien zu unterstützen.


|
|
|**Feature**|**Beschreibung**|
|:-----|:-----|
|Kontextbezogene Aktivierung|Outlook-Add-Ins sind kontextbezogen. Sie können basierend auf folgenden Kriterien aktiviert werden: 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>(Standard) für beliebige Elemente im Postfach oder Kalender</p></li><li><p>für einen bestimmten Elementtyp (eine E-Mail, eine Besprechungsanfrage oder einen Termin).</p></li><li><p>für eine Elementnachrichtenklasse.</p></li><li><p>für bestimmte Entitäten in einem Nachricht oder einem Termin. Siehe <span sdata="link"><a href="2cd5d8f1-69b3-4a2a-b31e-81a07a7cdd9f.htm">Kontextbezogene Outlook-Add-Ins</a></span>. </p></li><li><p>basierend auf bestimmten Regeln oder regulären Ausdrücken. Siehe <span sdata="link"><a href="b3fd6d69-b968-461d-a40e-6063f4febfe6.htm">Aktivierungsregeln für Outlook-Add-Ins</a></span> und <span sdata="link"><a href="93504f92-896f-4c80-9205-ba0b125f4290.htm">Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins</a></span>. </p></li><li><p>für Zeichenfolgenübereinstimmungen von Eigenschaften. Siehe <span sdata="link"><a href="a6b0904b-afe9-4882-9136-3d8cfd57fcf8.htm">Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten</a></span>.</p></li></ul>|
|Add-In-Befehle|Outlook-Add-In-Befehle ermöglichen die Initiierung von Add-In-Aktionen über das Menüband. Sie sind nur verfügbar für Add-Ins, die für alle E-Mails oder Ereignisse gelten. Weitere Informationen finden Sie unter [Add-In-Befehle für Outlook](a806cdfa-4230-4bcb-bb3f-7e3d1c2f26c2.md). |
|Roamingeinstellungen|Ein Outlook-Add-In kann Daten speichern, die spezifisch für das Benutzerpostfach sind, auf das Sie in einer späteren Outlook-Sitzung zugreifen können. Weitere Informationen finden Sie unter [Abrufen und Festlegen von Add-In-Metadaten für ein Outlook-Add-In](f10af299-2690-4714-8c7c-5c3d5157ef0c.md). |
|Benutzerdefinierte Eigenschaften|Ein Outlook-Add-In kann Daten speichern, die spezifisch für das Postfach des Benutzers sind, auf das Sie in einer späteren Outlook-Sitzung zugreifen können. Weitere Informationen finden Sie unter [Abrufen und Festlegen von Add-In-Metadaten für ein Outlook-Add-In](f10af299-2690-4714-8c7c-5c3d5157ef0c.md).|
|Abrufen von Anlagen oder des gesamten ausgewählten Elements|Ein Outlook-Add-In kann von der Serverseite auf Anlagen und das gesamte ausgewählte Element zugreifen. Weitere Informationen finden Sie hier:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Anlagen – Siehe <span sdata="link"><a href="0f872924-ea1a-4aa2-bb7b-e12d31014612.htm">Abrufen von Anlagen eines Outlook-Elements vom Server</a></span> und <span sdata="link"><a href="62669c4d-6829-4476-bac2-cac95fc0961e.htm">Hinzufügen und Entfernen von Anhängen bei einem Element in einem Formular zum Verfassen in Outlook</a></span>.</p></li><li><p>Gesamtes ausgewähltes Element – Dies gleicht dem Verwenden eines Rückruftokens zum Abrufen von Anlagen. Weitere Informationen finden Sie hier:</p><ul><li><p><a href="../reference/outlook/Office.context.mailbox.html(Office.15).aspx#getCallbackTokenAsync" target="_blank">mailbox.getCallbackTokenAsync</a>-Methode. Stellt ein Rückruftoken bereit, um den serverseitigen Code des Add-Ins für den Exchange Server zu identifizieren.</p></li><li><p><a href="https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#itemId" target="_blank">item.itemId </a>-Eigenschaft. Identifiziert das Element, das der Benutzer liest und das der serverseitige Code abruft.</p></li><li><p><a href="https://dev.outlook.com/reference/add-ins/Office.context.mailbox.md(Office.15).aspx#ewsUrl" target="_blank">mailbox.ewsUrl </a>-. Stellt die EWS-Endpunkt-URL bereit, die der serverseitige Code, zusammen mit dem Rückruftoken und der Element-ID, verwenden kann, um auf den EWS-Vorgang <a href="http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4(Office.15).aspx" target="_blank">GetItem </a> zuzugreifen, um das gesamte Element abzurufen.</p></li></ul></li></ul>|
|Benutzerprofil|Ein Mail-Add-In kann auf den Anzeigenamen, die E-Mail-Adresse und die Zeitzone in einem Benutzerprofil zugreifen. Weitere Informationen finden Sie unter dem [UserProfile](../reference/outlook/Office.context.mailbox.userProfile.md%28Office.15%29.md)-Objekt.|

## Erste Schritte zum Erstellen von Outlook-Add-Ins

Weitere Informationen zu den ersten Schritten beim Erstellen von Outlook-Add-Ins finden Sie unter [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx).


## Zusätzliche Ressourcen


Andere allgemeine für die Entwicklung von Office-Add-Ins geltende Konzepte finden Sie unter:


- [Designrichtlinien für Office-Add-Ins](d5b2ab2e-dfc8-47c8-919c-e9c23358d70c.md)
    
- [Bewährte Methoden für die Entwicklung von Office-Add-Ins](013e1486-4482-42c1-bcda-edf8de06e771.md)
    
- [Lizenzieren von Office- und SharePoint-Add-Ins](http://msdn.microsoft.com/library/3e0e8ff6-66d6-44ff-b0c2-59108ebd9181%28Office.15%29.aspx)
    
- [Veröffentlichen von Office- und SharePoint-Add-Ins und Office 365 Web Apps im Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
- [JavaScript-API für Office](http://msdn.microsoft.com/EN-US/library/fp142185%28v=office.15%29.aspx(Office.15).aspx)
    
- [Mail-Add-In-Manifeste](98ee2cca-f866-4377-bd1f-6a5338ebd7b1.md)
    
