
# Vergleichen der Outlook-Add-In-Unterstützung in Outlook für Mac mit anderen Outlook-Hosts
Die Outlook-Add-In-Unterstützung in Outlook für Mac unterscheidet sich von der in anderen unterstützten Outlook-Hosts.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Sie können ein Outlook-Add-In in Outlook für Mac auf die gleiche Weise erstellen und ausführen wie in den anderen Hosts, darunter Outlook für Windows, OWA für mobile Geräte und Outlook Web App, ohne dass Sie den JavaScript-Code für die einzelnen Hosts anpassen müssen. Ein Aufruf eines Add-Ins an die JavaScript-API für Office funktioniert in der Regel stets gleich, mit Ausnahme der in der folgenden Tabelle beschriebenen Bereiche.

(../../images  Outlook für Mac unterstützt JavaScript-API für Office nur im Outlook-Lesemodus.



|**Bereich**|**Outlook für Windows, OWA für mobile Geräte, Outlook Web App**|**Outlook für Mac**|
|:-----|:-----|:-----|
|Unterstützte Versionen von office.js und Office-Add-Ins-Manifestschema|Alle APIs in Office.js und Schema v1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Nur APIs für den Lesemodus. Ein Add-In, das die <a href="72915b13-720f-4dc5-b5d1-4676e2a536ba.htm#mod_off15_WhatsNewMailApps_NewJSAPI">neuen und erweiterten APIs in office.js v1.1</a> verwendet, kann zwar aktiviert werden, die APIs für den Verfassenmodus werden jedoch in Outlook für Mac nicht ordnungsgemäß ausgeführt. </p></li><li><p>Schema v1.1</p></li></ul>|
|Instanzen einer Terminserie|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Die Element-ID und andere Eigenschaften eines Mastertermins oder einer Termininstanz einer Serie können abgerufen werden. </p></li><li><p>Mit <a href="../reference/outlook/Office.context.mailbox.md(Office.15).aspx#displayAppointmentForm" target="_blank">mailbox.displayAppointmentForm</a> kann eine Instanz oder der Master einer Terminserie angezeigt werden.</p></li></ul>|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Die Element-ID und andere Eigenschaften des Mastertermins können abgerufen werden, nicht jedoch die Werte einer Instanz einer Terminserie.</p></li><li><p>Der Mastertermin einer Terminserie kann angezeigt werden. Ohne Element-ID kann keine Instanz einer Terminserie angezeigt werden.</p></li></ul>|
|Empfängertyp eines Terminteilnehmers|Mit [EmailAddressDetails.recipientType](../reference/outlook/simple-types.md%28Office.15%29.md) kann der Empfängertyp eines Teilnehmers identifiziert werden.|**EmailAddressDetails.recipientType** gibt **undefined** für Terminteilnehmer zurück.|
|Versionszeichenfolge des Hosts |Das Format der von [diagnostics.hostVersion](../reference/outlook/Office.context.mailbox.diagnostics.md%28Office.15%29.md) zurückgegebenen Versionszeichenfolge hängt von tatsächlichen Typ des Hosts ab. Beispiel:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Outlook für Windows: 15.0.4454.1002</p></li><li><p>Outlook Web App: 15.0.918.2</p></li></ul>|Beispiel für die von  **Diagnostics.hostVersion** zurückgegebene Versionszeichenfolge in Outlook für Mac: 15.0 (140325)|
|Benutzerdefinierte Eigenschaften eines Elements|Wenn das Netzwerk ausfällt, kann ein Add-In weiterhin auf zwischengespeicherte benutzerdefinierte Eigenschaften zugreifen.|Da Outlook für Mac benutzerdefinierte Eigenschaften nicht zwischenspeichert, können Add-Ins bei einem Netzwerkausfall nicht auf diese Eigenschaften zugreifen.|
|Anlagendetails|Der Inhaltstyp und die Anlagennamen in einem [AttachmentDetails](https://dev.outlook.com/reference/add-ins/simple-types.mdl.aspx#AttachmentDetails)-Objekt hängen vom Hosttyp ab:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>JSON-Beispiel für <span class="keyword">AttachmentDetails.contentType</span>: <span class="keyword">"contentType": "image/x-png"</span>. </p></li><li><p><span class="keyword">AttachmentDetails.name</span> enthält keine Dateierweiterung. Wenn die Anlage zum Beispiel eine Nachricht mit dem Betreff "RE: Summer activity" ist, lautet das JSON-Objekt, das diesen Anlagennamen darstellt: <span class="keyword">"name": "RE: Summer activity"</span>.</p></li></ul>|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>JSON-Beispiel für <span class="keyword">AttachmentDetails.contentType</span>: <span class="keyword">"contentType": "image/png"</span></p></li><li><p><span class="keyword">AttachmentDetails.name</span> enthält immer eine Dateierweiterung. Anlagen in Form von E-Mail-Elementen haben die Erweiterung ".eml", und Termine haben die Erweiterung ".ics". Wenn die Anlage beispielsweise eine E-Mail mit dem Betreff "RE: Summer activity" ist, sieht das JSON-Objekt, das den Anlagennamen darstellt, folgendermaßen aus: <span class="keyword">"name": "RE: Summer activity.eml"</span>.</p></li></ul>|
|Zeichenfolge zur Darstellung der Zeitzone in den Eigenschaften  **dateTimeCreated** und **dateTimeModified**|Beispiel: Thu Mar 13 2014 14:09:11 GMT+0800 (China Normalzeit)|Beispiel: Thu Mar 13 2014 14:09:11 GMT+0800 (CST)|
|Zeitgenauigkeit von  **dateTimeCreated** und **dateTimeModified**|Wenn ein Add-In den folgenden Code verwendet, ist die Angabe auf die Millisekunde genau.
```
JSON.stringify(Office.context.mailbox.item, null, 4);

```

|Ist die Angabe nur auf die Sekunde genau.|

## Zusätzliche Ressourcen



- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md)
    
