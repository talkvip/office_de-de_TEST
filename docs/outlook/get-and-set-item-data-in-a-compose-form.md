
# Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook
Informationen zum Abrufen oder Festlegen von Eigenschaften eines Outlook-Add-Ins in einem Erstellungsszenario, einschließlich Empfänger, Betreff, Text sowie Ort und Uhrzeit eines Termins.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Elementeigenschaften für ein Entwurfs-Add-in abrufen und festlegen


Sie können in einer Entwurfsvorlage die meisten der Eigenschaften abrufen, die in einem ähnlichen Element wie einem Leseformular eingeblendet sind (wie Teilnehmer, Empfänger, Betreff und Text). Sie können außerdem bestimmte Extraeigenschaften abrufen, die für eine Entwurfsvorlage nicht aber für ein Leseformular (Text, BCC) von Bedeutung sind. 

Bei den meisten dieser Eigenschaften sind die Methoden zum Abrufen und Festlegen der Eigenschaften asynchron, da es möglich ist, dass ein Outlook-Add-In und der Benutzer dieselbe Eigenschaft gleichzeitig ändern können. In Tabelle 1 werden die Eigenschaften auf Elementebene und die entsprechenden asynchronen Methoden zum Abrufen und Festlegen von Eigenschaften in einer Entwurfsvorlage angezeigt. Die Eigenschaften [item.itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [item.conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) sind Ausnahmen, da Benutzer sie nicht ändern dürfen. Sie können die Eigenschaften programmtechnisch auf dieselbe Weise über eine Entwurfsvorlage wie über ein Leseformular direkt aus dem übergeordneten Objekt abrufen.

Sie können nicht nur über die JavaScript-API für Office auf die Elementeigenschaften zugreifen, sondern auch über Exchange Web Services (EWS). Anhand der  **ReadWriteMailbox**-Berechtigung können Sie die [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode zum Zugriff auf EWS-Vorgänge und [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) verwenden, um weitere Eigenschaften von Elementen für das Benutzerpostfach abzurufen und festzulegen. **makeEwsRequestAsync** ist sowohl in Erstell- als auch in Leseformularen verfügbar. Weitere Informationen zur **ReadWriteMailbox**-Berechtigung und zum Zugriff auf EWS über die Office-Add-Ins-Plattform finden Sie unter [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../../docs/outlook/privacy/understanding-outlook-add-in-permissions.md) und [Aufrufen von Webdiensten aus einem Outlook-Add-In](../outlook/web-services.md).


**Tabelle 1. Asynchrone Methoden zum Abrufen und Festlegen von Elementeigenschaften in einer Entwurfsvorlage**


|**Eigenschaft**|**Eigenschaftentyp**|**Asynchrone Abrufmethode**|**Asynchrone Festlegungsmethode(n)**|
|:-----|:-----|:-----|:-----|
|[bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|[Recipients.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|[Recipients.addAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)[Recipients.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|
|[body](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Body](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|[Body.getAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|[Body.prependAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)[Body.setAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)[Body.setSelectedDataAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|
|[cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[end](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Time](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|[Time.getAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|[Time.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|
|[location](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|[Location.getAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|[Location.setAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|
|[optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[start](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Time**|**Time.getAsync**|**Time.setAsync**|
|[subject](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|[Subject.getAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|[Subject.setAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|
|[to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|

## Inhalt dieses Abschnitts:


In den folgenden Beispielen sehen Sie, wie Sie Elementeigenschaften asynchron abrufen oder festlegen können.


- [Abrufen, Festlegen oder Hinzufügen von Empfängern beim Erstellen eines Termins oder einer Nachricht in Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Abrufen oder Festlegen des Betreffs, wenn Sie einen Termin oder eine Nachricht in Outlook verfassen](../outlook/get-or-set-the-subject.md)
    
- [Einfügen von Daten in den Textkörper bei der Erstellung eines Termins oder einer Nachricht in Outlook](../outlook/insert-data-in-the-body.md)
    
- [Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen](../outlook/get-or-set-the-time-of-an-appointment.md)
    

## Zusätzliche Ressourcen



- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](../outlook/compose-scenario.md)
    
- [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../../docs/outlook/privacy/understanding-outlook-add-in-permissions.md)
    
- [Aufrufen von Webdiensten aus einem Outlook-Add-In](../outlook/web-services.md)
    
- [Abrufen und Festlegen von Outlook-Elementdaten in Formularen zum Lesen oder Verfassen](../outlook/item-data.md)
    


