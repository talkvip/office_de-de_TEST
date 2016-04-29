
# Abrufen und Festlegen von Outlook-Elementdaten in Formularen zum Lesen oder Verfassen
Informieren Sie sich über die wenigen Möglichkeiten zum Abrufen oder Festlegen von Daten für einen Termin oder eine Nachricht von einem Outlook-Add-In, wobei berücksichtigt wird, ob der Benutzer das Element anzeigt oder erstellt. Sie können Elementebeneneigenschaften in der JavaScript-API für Office, im Exchange Server Rückruftoken oder in Exchange-Webdiensten verwenden.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## In Lese- und Erstellformularen verfügbare Eigenschaften


Seit Version 1.1 des Office-Add-Ins-Manifestschemas kann Outlook Add-Ins aktivieren, wenn der Benutzer ein Element anzeigt oder erstellt. Je nachdem, ob ein Add-In in einem Lese- oder Erstellformular aktiviert ist, unterscheiden sich auch die Eigenschaften, die für das Add-In in dem Element verfügbar sind. Die Eigenschaften [dateTimeCreated](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [dateTimeModified](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) sind beispielsweise nur für ein Element definiert, das bereits gesendet wurde (das Element wird anschließend in einem Leseformular angezeigt), nicht jedoch beim Erstellen des Elements (in einem Erstellformular). Ein weiteres Beispiel ist die Eigenschaft [bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md). Sie ist nur beim Erstellen einer Nachricht (in einem Erstellformular) aussagekräftig und steht dem Benutzer in einem Leseformular nicht zur Verfügung.

Tabelle 1 enthält die Elementebeneneigenschaften in der JavaScript-API für Office, die im Lese- und Verfassenmodus der Mail-Add-Ins zur Verfügung stehen. In der Regel sind die in Leseformularen zur Verfügung gestellten Eigenschaften schreibgeschützt und die in Erstellformularen zur Verfügung gestellten Eigenschaften weisen Lese-/Schreibzugriff auf, mit Ausnahme der Eigenschaften [itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md), die immer schreibgeschützt sind. Bei den restlichen in Erstellformularen verfügbaren Elementebeneneigenschaften sind die Methoden zum Abrufen und Festlegen im Verfassenmodus asynchron, da das Add-In und der Benutzer die gleiche Eigenschaft möglicherweise gleichzeitig lesen oder schreiben, also unterscheidet sich der von diesen Eigenschaften zurückgegebene Objekttyp in Erstell- und Leseformularen ebenfalls. Weitere Informationen zum Verwenden asynchroner Methoden zum Abrufen oder Festlegen von Elementebeneneigenschaften im Verfassenmodus finden Sie unter [Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook](60901b90-2472-4132-9b5e-0749e2e33cbe.md).


**Tabelle 1. In Erstell- und Leseformularen verfügbare Eigenschaften**


|**Elementtyp**|**Eigenschaft**|**Eigenschaftstyp in Leseformularen**|**Eigenschaftstyp in Erstellformularen**|
|:-----|:-----|:-----|:-----|
|Termine und Nachrichten|[dateTimeCreated](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Date**-Objekt (JavaScript)|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[dateTimeModified](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Date**-Objekt (JavaScript)|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[itemClass](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|Zeichenfolge in [ItemType](http://dev.outlook.com/reference/add-ins/Office.MailboxEnums.html%28Office.15%29.md)-Aufzählung|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[Anlagen](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[AttachmentDetails](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|Eigenschaft nicht verfügbar|
|Termine und Nachrichten|[Text](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Text](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|[Text](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|
|Termine|[end](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Date**-Objekt (JavaScript)|[Zeit](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|
|Termine|[Ort](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|[Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|
|Termine und Nachrichten|[normalizedSubject](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|Eigenschaft nicht verfügbar|
|Termine|[optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[EmailAddressDetails](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)|[Empfänger](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|
|Termine|[Organisator](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|Eigenschaft nicht verfügbar|
|Termine|[requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|**Recipients**|
|Termine|[Ressourcen](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|Eigenschaft nicht verfügbar|
|Termine|[start](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Date**-Objekt (JavaScript)|**Time**|
|Termine und Nachrichten|[Betreff](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|
|Nachrichten|[bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|Eigenschaft nicht verfügbar|**Recipients**|
|Nachrichten|[cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|**Recipients**|
|Meldungen|[conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|String|Zeichenfolge (schreibgeschützt)|
|Meldungen|[from](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|Eigenschaft nicht verfügbar|
|Nachrichten|[internetMessageId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|Integer|Eigenschaft nicht verfügbar|
|Nachrichten|[Absender](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|Eigenschaft nicht verfügbar|
|Nachrichten|[in](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**EmailAddressDetails**|**Recipients**|

## Verwenden von Exchange Server Rückruftoken aus einem Lese-Add-in


Wenn Ihr Outlook-Add-In in Leseformularen aktiviert ist, können Sie ein Exchange-Rückruftoken abrufen. Dieses Token kann in serverseitigem Code zum Zugreifen auf das vollständige Element über Exchange Web Services (EWS) verwendet werden. Indem Sie die  **ReadItem**-Berechtigung im Add-In-Manifest angeben, können Sie die Methode [mailbox.getCallbackTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md) um Abrufen eines Exchange-Rückruftokens, die Eigenschaft [mailbox.ewsUrl](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md) zum Abrufen der URL des EWS-Endpunkts für das Postfach des Benutzers und [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) zum Abrufen der EWS-ID des ausgewählten Elements verwenden. Anschließend können Sie das Rückruftoken, die EWS-Endpunkt-URL und die EWS-Element-ID an den serverseitigen Code weitergeben, um auf den [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) -Vorgang zum Abrufen weiterer Eigenschaften des Elements zuzugreifen.


## Zugreifen auf Exchange-Webdienste aus einem Lese- oder Erstellungs-Add-in


Sie können auch die Methode [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)zum Zugreifen auf EWS-Vorgänge sowie [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) und [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) direkt aus dem Add-In verwenden. Sie können diese Eigenschaften zum Abrufen und Festlegen vieler Eigenschaften eines bestimmten Elements verwenden. Diese Methode und die EWS-Vorgänge sind in den Outlook-Add-Ins unabhängig davon verfügbar, ob das Add-In in einem Lese- oder Erstellformular aktiviert wurde, vorausgesetzt Sie haben die **ReadWriteMailbox**-Berechtigung im Add-In-Manifest angegeben. Weitere Informationen zum Verwenden von  **makeEwsRequestAsync** für den Zugriff auf EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-Add-In](0049645e-9af1-46c7-a22e-fe6c650b134e.md).


## Zusätzliche Ressourcen



- [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md)
    
- [Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook](60901b90-2472-4132-9b5e-0749e2e33cbe.md)
    
- [Aufrufen von Webdiensten aus einem Outlook-Add-In](0049645e-9af1-46c7-a22e-fe6c650b134e.md)
    


