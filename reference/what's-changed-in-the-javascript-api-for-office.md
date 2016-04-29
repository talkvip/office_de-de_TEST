
# Änderungen in der JavaScript-API für Office
Die JavaScript-API für Office bietet neue und aktualisierte Objekte, Methoden, Eigenschaften, Ereignisse und Enumerationen, welche den Funktionsumfang Ihrer Office-Add-Ins erweitern. Verwenden Sie die Links unten, um die neuen und aktualisierten API-Mitglieder anzuzeigen.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Um Add-ins mithilfe der neuen API-Mitglieder zu entwickeln, müssen Sie [die Dateien der JavaScript-API für Office in Ihrem Projekt aktualisieren](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Alle API-Mitglieder, einschließlich derjenigen, die aus vorherigen Updates unverändert übernommen wurden, finden Sie unter [JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx).


## Neue und aktualisierte API

 **Neue und aktualisierte Objekte**



|**Objekt**|**Beschreibung**|**Version hinzugefügt oder aktualisiert **|
|:-----|:-----|:-----|
|[Item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Die Methoden <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> und <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> wurden eingeführt, um die Benutzerauswahl abrufen und im Betreff oder Körper der Nachricht oder des Termins zu überschreiben.</p></li><li><p>Die Methoden <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#displayReplyAllForm" target="_blank">displayReplyAllForm</a> und <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#displayReplyForm" target="_blank">displayReplyForm</a> wurden aktualisiert, damit Anhänge an ein Antwortformular zu einem Termin angehängt werden können.</p></li></ul>|Mailbox 1.2|
|[Item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|Mit Methoden und Feldern zum Erstellen von Verfassenmodus-Outlook-Add-Ins aktualisiert. |1.1|
|[Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx)|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|[Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|[Body](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|Wurde hinzugefügt, um das Erstellen und Bearbeiten des Textkörpers einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)|Aktualisierungen und Ergänzungen:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Unterstützung der Eigenschaften <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> und <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> in Inhalts-Add-ins für Access</p></li><li><p>Abrufen des Dokuments als PDF mit der Methode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> in Add-ins für PowerPoint und Word</p></li><li><p>Abrufen von Dateieigenschaften mit der Methode <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> in Add-ins für Excel, PowerPoint und Word</p></li><li><p>Navigieren zu Speicherorten und Objekten im Dokument mit der Methode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> in Add-ins für Excel und PowerPoint</p></li><li><p>Abrufen der ID, des Titels und des Indexes für ausgewählte Folien mit der Methode <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (wenn Sie die neue <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>-Enumeration angeben) in Add-ins für PowerPoint</p></li></ul>|1.1|
|[Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|Wurde hinzugefügt, um das Festlegen des Orts eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx)|Die Auswahlmethode unterstützt jetzt das Abrufen von Bindungen in Inhalts-Add-ins für Access.|1.1|
|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|Wurde hinzugefügt, um das Abrufen und Festlegen der Empfänger einer Nachricht oder eines Termins im Verfassenmodus zu ermöglichen.|1.1|
|[Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)|Unterstützt jetzt das Erstellen benutzerdefinierter Einstellungen in Inhalts-Add-ins für Access.|1.1|
|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|Wurde hinzugefügt, um das Abrufen und Festlegen des Betreffs einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Time](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|Wurde hinzugefügt, um das Abrufen und Festlegen der Start- und Endzeit eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|

**Neue und aktualisierte Enumerationen**


|**Objekt**|**Beschreibung**|**Version**|
|:-----|:-----|:-----|
|[ActiveView](http://msdn.microsoft.com/library/1f1d963e-04e1-4cf2-b161-5329d7ad0a3e%28Office.15%29.aspx)|Gibt den Status der aktiven Ansicht des Dokuments an, z. B. ob der Benutzer das Dokument bearbeiten kann.Wurde hinzugefügt, damit Add-ins für PowerPoint ermitteln können, ob die Benutzer die Präsentation ( **Diaschau**) anzeigen oder Folien bearbeiten. |1.1|
|[CoercionType](http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx)|Mit  **Office.CoercionType.SlideRange** aktualisiert, um das Abrufen des ausgewählten Folienbereichs mit der Methode **getSelectedDataAsync** in Add-ins für PowerPoint zu unterstützen.|1.1|
|[EventType](http://msdn.microsoft.com/library/82c79659-52da-48b0-92a9-831226eb9a7f%28Office.15%29.aspx)|Mit dem neuen ActiveViewChanged-Ereignis aktualisiert.|1.1|
|[FileType](http://msdn.microsoft.com/library/fadbb4cf-a0e4-47b2-93dd-123f0b06d4ae%28Office.15%29.aspx)|Gibt jetzt Ausgaben in PDF-Format an.|1.1|
|[GoToType](http://msdn.microsoft.com/library/8de45be3-de35-4765-a67a-e128a46786bd%28Office.15%29.aspx)|Wurde hinzugefügt, um die Stelle oder das Objekt im Dokument anzugeben, zu dem gewechselt werden soll.|1.1|

## Zusätzliche Ressourcen


- [API und Schemaverweise für Office-Add-Ins](../../reference/reference.md)
    
- [Office-Add-Ins](../../docs/overview/office-add-ins.md)
    
