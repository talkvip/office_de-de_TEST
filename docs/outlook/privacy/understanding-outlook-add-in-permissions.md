
# Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer
Verwenden Sie das Berechtigungsmodell für Outlook-Add-Ins zum Anfordern des geeigneten Postfachzugriffs für ein Add-In:  **Restricted**,  **ReadItem**,  **ReadWriteItem** oder **ReadWriteMailbox**.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Berechtigungsmodell für Outlook-Add-Ins


Outlook-Add-Ins geben die erforderliche Berechtigungsstufe in ihrem Manifest an. Die verfügbaren Stufen sind  **Restricted**,  **ReadItem**,  **ReadWriteItem** oder **ReadWriteMailbox**. Diese Berechtigungsstufen sind kumulativ:  **Restricted** ist die niedrigste Ebene, und jede höhere Ebene enthält die Berechtigungen aller niedrigeren Ebenen. **ReadWriteMailbox** enthält alle unterstützten Berechtigungen.

Sie können die von einem Mail-Add-In angeforderten Berechtigungen einsehen, bevor Sie es aus dem Office Store installieren. Die für installierte Add-Ins erforderlichen Berechtigungen können Sie auch im Exchange Admin Center anzeigen.


## Berechtigung "Eingeschränkt"


Die Berechtigung  **Restricted** ist die grundlegendste Berechtigungsstufe. Geben Sie **Restricted** im [Permissions](http://msdn.microsoft.com/de-de/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)-Element im Manifest an, um diese Berechtigung anzufordern. Outlook weist diese Berechtigung eines Mail-Add-Ins standardmäßig zu, wenn das Add-In in seiner Manifestdatei keine bestimmte Berechtigung anfordert.


### Möglich


- Nur bestimmte Entitäten (Telefonnummer, Adresse, URL) aus dem Betreff oder Textkörper des Elements abrufen.
    
- Eine [ItemIs-Aktivierungsregel](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemIs) angeben, bei der das aktuelle Element in einem Lese- oder Erstellformular ein bestimmter Elementtyp sein muss oder eine [ItemHasKnownEntity-Regel](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Activating) angeben, die mit einem der kleineren Teile der unterstützten bekannten Entitäten (Telefonnummer, Adresse, URL) im ausgewählten Element übereinstimmen.
    
- Auf Eigenschaften und Methoden zugreifen, die  **nicht** für bestimmte Informationen über den Benutzer oder das Element gelten. (Im nächsten Abschnitt finden Sie eine Liste der Elemente, bei denen sie gelten).
    

### Nicht möglich


- Verwenden einer [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regel für den Vertrag, die E-Mail-Adresse, den Besprechungsvorschlag oder die Vorgangsvorschlagentität. 
    
- Verwenden der [ItemHasAttachment](http://msdn.microsoft.com/de-de/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)- oder [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regel. 
    
- Zugreifen auf die Elemente der folgenden Liste, die für die Informationen über den Benutzer oder das Element gelten. Beim Versuch des Zugriffs auf Elemente in dieser Liste werden  **null** und die Fehlermeldung zurückgegeben, dass Outlook für das Mail-Add-in erhöhte Berechtigungen anfordert.
    
      - [item.addFileAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.addItemAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.attachments](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.body](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.from](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.getRegExMatches](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.getRegExMatchesByName](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.organizer](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.removeAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.resources](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.sender](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [item.to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
  - [mailbox.getCallbackTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)
    
  - [mailbox.getUserIdentityTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)
    
  - [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)
    
  - [mailbox.userProfile](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.userProfile.html%28Office.15%29.md)
    
  - [Body](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md) und alle untergeordneten Member
    
  - [Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md) und alle untergeordneten Member
    
  - [Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md) und alle untergeordneten Member
    
  - [Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md) und alle untergeordneten Member
    
  - [Time](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md) und alle untergeordneten Member
    

## ReadItem-Berechtigung


Die Berechtigung  **ReadItem** ist die nächste Stufe im Berechtigungsmodell. Geben Sie **ReadItem** im **Permissions**-Element in der Manifestdatei an, um diese Berechtigung anzufordern. 


### Möglich


- [Lesen aller Eigenschaften](../outlook/item-data.md) des aktuellen Elements in einem Lese- oder [Erstellformular](../outlook/get-and-set-item-data-in-a-compose-form.md), z. B. [item.to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) in einem Leseformular und [item.to.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md) in einem Erstellformular.
    
- [Abrufen eines Rückruftoken zum Abrufen der Elementanlagen](../outlook/get-attachments-of-an-outlook-item.md) oder des vollständigen Elements.
    
- [Schreiben von benutzerdefinierten Eigenschaften](30217d63-7615-4f3f-8618-c91e4e60cd43.md), die von dem Add-in für das Element festgelegt wurden.
    
- [Abrufen aller vorhandenen bekannten Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Retrieving), statt nur einen Teil des Betreff oder Textkörpers des Elements.
    
- Verwenden Sie alle [bekannten Entitäten](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemHasKnownEntity) in [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regeln oder [regulären Ausdrücken](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemHasRegularExpressionMatch) in [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regeln. Das folgende Beispiel entspricht Schema v1.1. Es zeigt eine Regel, die das Add-in aktiviert, wenn mindestens eine bekannte Entität im Betreff oder Textkörper der ausgewählten Nachricht gefunden wird:
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### Nicht möglich

Greifen Sie auf  **mailbox.makeEWSRequestAsync** oder die folgenden Write-Methoden zu:


- [item.addFileAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
- [item.addItemAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
- [item.bcc.addAsync](../reference/outlook/Recipients.md%28Office.15%29.md)
    
- [item.bcc.setAsync](../reference/outlook/Recipients.md%28Office.15%29.md)
    
- [item.body.prependAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)
    
- [item.body.setAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)
    
- [item.body.setSelectedDataAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)
    
- [item.cc.addAsync](../reference/outlook/Recipients.md%28Office.15%29.md)
    
- [item.cc.setAsync](../reference/outlook/Recipients.md%28Office.15%29.md)
    
- [item.end.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)
    
- [item.location.setAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)
    
- [item.optionalAttendees.addAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    
- [item.optionalAttendees.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    
- [item.removeAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)
    
- [item.requiredAttendees.addAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    
- [item.requiredAttendees.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    
- [item.start.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)
    
- [item.subject.setAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)
    
- [item.to.addAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    
- [item.to.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)
    

## ReadWriteItem-Berechtigung


Geben Sie die Berechtigung  **ReadWriteItem** im **Permissions**-Element in der Manifestdatei an, um diese Berechtigung anzufordern. In Erstellformularen aktivierte Mail-Add-ins, die Schreibmethoden ( **Message.to.addAsync** oder **Message.to.setAsync**) verwenden, benötigen mindestens diese Berechtigungsstufe.


### Möglich


- [Lesen und schreiben aller Elementebeneneigenschaften](../outlook/item-data.md) des Elements, das in Outlook angezeigt oder verfasst wird.
    
- [Hinzufügen oder entfernen von Anlagen](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md) des Elements.
    
- Verwenden Sie alle anderen Elemente der JavaScript-API für Office, die für Mail-Add-ins anwendbar sind, außer  **Mailbox.makeEWSRequestAsync**.
    

### Nicht möglich

Verwenden der Methode  **Mailbox.makeEWSRequestAsync**.


## ReadWriteMailbox-Berechtigung


Die Berechtigung  **ReadWriteMailbox** ist die höchste Berechtigungsstufe. Geben Sie die Berechtigung **ReadWriteMailbox** im **Permissions**-Element in der Manifestdatei an, um diese Berechtigung anzufordern. 

Zusätzlich zu den unterstützten Funktionen der Berechtigung  **ReadWriteItem** durch die Verwendung der Methode **Mailbox.makeEWSRequestAsync**, können Sie auf unterstützte Vorgänge des Exchange-Webdiensts zugreifen, um Folgende Aktionen auszuführen:


- Lesen und schreiben aller Eigenschaften eines beliebigen Elements im Postfach des Benutzers.
    
- Erstellen, lesen und schreiben eines beliebigen Ordners oder Elements in dem Postfach.
    
- Senden eines Elements aus dem Postfach
    
Über die Methode  **mailbox.makeEWSRequestAsync** können Sie auf folgende EWS-Vorgänge zugreifen:


- [CopyItem](http://msdn.microsoft.com/de-de/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/de-de/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/de-de/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/de-de/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/de-de/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/de-de/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/de-de/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/de-de/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/de-de/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/de-de/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/de-de/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/de-de/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/de-de/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/de-de/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
Der Versuch, einen nicht unterstützten Vorgang auszuführen, führt zu einer Fehlermeldung.


## Zusätzliche Ressourcen



- [Datenschutz, Berechtigungen und Sicherheit für Outlook-Add-Ins](../../docs/develop/privacy-and-security.md)
    
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
