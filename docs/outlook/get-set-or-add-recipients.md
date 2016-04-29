
# Abrufen, Festlegen oder Hinzufügen von Empfängern beim Erstellen eines Termins oder einer Nachricht in Outlook
Erfahren Sie mehr darüber, wie ein Outlook-Add-In Empfänger abrufen, festlegen oder hinzufügen kann, wenn ein Benutzer einen Termin oder eine Nachricht in Outlook erstellt.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Voraussetzungen für das Abrufen, Festlegen und Hinzufügen von Empfängern in einem Formular zum Verfassen


Die JavaScript-API für Office stellt asynchrone Methoden bereit ([Recipients.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md), [Recipients.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md) oder [Recipients.addAysnc](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)) um Empfänger in einem Formular zum Verfassen für einen Termin oder eine Nachricht abzurufen, festzulegen oder hinzuzufügen. Diese asynchronen Methoden sind ausschließlich für Verfassen-Add-Ins verfügbar. Um diese Methoden verwenden zu können, müssen Sie sicherstellen, dass Sie das Add-In-Manifest richtig eingerichtet haben, sodass Outlook das Add-In in Formularen zum Verfassen aktiviert, wie im Abschnitt [Einrichten von Outlook-Add-Ins für Erstellformulare](e4126e58-4ddc-4891-9f19-aa6c1a258027.md#mod_off15_CreatingForCompose_SettingUp) von [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md) beschrieben.

Einige der Eigenschaften, die Empfänger in einem Termin oder einer Nachricht darstellen, sind für den Lesezugriff in einem Formular zum Verfassen und in einem Formular zum Lesen verfügbar. Diese Eigenschaften sind u. a. [optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) für Termine sowie [cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) für Nachrichten. In Formularen zum Lesen können Sie direkt über das übergeordnete Objekt auf die Eigenschaft zugreifen:




```
item.cc
```

Aber weil in einem Formular zum Verfassen sowohl der Benutzer als auch Ihr Add-In gleichzeitig einen Empfänger einfügen oder ändern können, müssen Sie die asynchrone Methode  **getAsync** verwenden, um diese Eigenschaften zu erhalten:




```
item.cc.getAsync
```

Diese Eigenschaften sind nur für den Schreibzugriff in Formularen zum Verfassen und zum Lesen verfügbar.

Wie die meisten asynchronen Methoden in der JavaScript-API für Office verwenden  **getAsync**,  **setAsync** und **addAsync** optionale Eingabeparameter. Weitere Informationen für das Angeben dieser optionalen Eingabeparameter finden Sie unter [Übergeben optionaler Parameter an asynchrone Methoden](http://msdn.microsoft.com/de-de/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters) in [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md).


## So werden Empfänger abgerufen


In diesem Abschnitt wird ein Codebeispiel beschrieben, über das die Empfänger eines Termins oder einer Nachricht, der/die verfasst wird, abgerufen und die E-Mail-Adressen der Empfänger angezeigt werden. Das Codebeispiel setzt eine Regel im Add-in-Manifest voraus, die das Add-In in einem Formular zum Verfassen für einen Termin oder eine Nachricht aktiviert: 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Weil sich die Eigenschaften, die die Empfänger eines Termins darstellen ( **optionalAttendees** und **requiredAttendees**) von den Eigenschaften einer Nachricht ([bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md),  **cc** und **to**) unterscheiden, sollten Sie in der JavaScript-API für Office zunächst die [item.itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft verwenden, um festzustellen, ob das Element, das gerade verfasst wird, ein Termin oder eine Nachricht ist. Im Verfassenmodus sind diese Eigenschaften von Terminen und Nachrichten [Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Objekte, sodass Sie die asynchrone Methode,  **Recipients.getAsync**, anwenden können, um die einhergehenden Empfänger abzurufen. 

Um  **getAsync** verwenden zu können, müssen Sie eine Rückrufmethode bereitstellen, um Status, Ergebnisse und Fehler zu prüfen, die vom asynchronen **getAsync**-Abruf zurückgegeben werden. Mithilfe des optionalen  _asyncContext_-Parameters können Sie jedes beliebige Argument für die Rückrufmethode bereitstellen. Die Rückrufmethode gibt einen  _asyncResult_-Ausgabeparameter zurück. Sie können die  **status**- und  **error**-Eigenschaften des [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Parameterobjekts verwenden, um den Status des asynchronen Abrufs sowie auf Fehlernachrichten zu prüfen, die der Abruf möglicherweise ausgegeben hat, und Sie können die  **value**-Eigenschaft verwenden, um die tatsächlichen Empfänger abzurufen. Empfänger werden als ein Array von [EmailAddressDetails](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekten dargestellt.

Beachten Sie, dass aufgrund der Tatsache, dass die  **getAsync**-Methode asynchron ist, Sie Ihren Code bei Folgeaktionen, die von einem erfolgreichen Abruf der Empfänger abhängen, so organisieren sollten, dass solche Aktionen nur in der zugehörigen Rückrufmethode gestartet werden, nachdem der asynchrone Abruf erfolgreich beendet wurde. 




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## So werden Empfänger festgelegt


In diesem Abschnitt wird ein Codebeispiel beschrieben, das die Empfänger eines Termins oder einer Nachricht festlegt, der/die gerade vom Benutzer verfasst wird. Durch das Festlegen von Empfängern werden alle bereits bestehenden Empfänger überschrieben. Ähnlich wie im vorherigen Beispiel, in dem Empfänger in einem Formular zum Verfassen abgerufen werden, setzt dieses Beispiel voraus, dass das Add-In in Formularen zum Verfassen für Termine oder Nachrichten aktiviert wird. Dieses Beispiel verifiziert zunächst, ob das Element, das gerade verfasst wird, ein Termin oder eine Nachricht ist, sodass die asynchrone Methode  **Recipients.setAsync** auf die entsprechenden Eigenschaften angewendet wird, die die Empfänger des Termins oder der Nachricht darstellen.

Stellen Sie beim Abrufen von  **setAsync** ein Array als Eingabeargument für den _recipients_-Parameter in einem der folgenden Formate bereit:


- Ein Array von Zeichenfolgen, die SMTP-Adressen sind.
    
- Ein Array von Wörterbüchern, die jeweils einen Anzeigenamen und eine E-Mail-Adresse enthalten, wie im folgenden Beispiel dargestellt.
    
- Ein Array von  **EmailAddressDetails**-Objekten, ähnlich dem von der  **getAsync**-Methode zurückgegebenen Objekt.
    
Alternativ können Sie eine Rückrufmethode als Eingabeargument für die  **setAsync**-Methode bereitstellen, um sicherzustellen, dass der Code, der von der erfolgreichen Festlegung der Empfänger abhängt, nur dann ausgeführt wird, wenn dies geschieht. Sie können außerdem jedes beliebige Argument für die Rückrufmethode bereitstellen, indem Sie den optionalen  _asyncContext_-Parameter verwenden. Wenn Sie eine Rückrufmethode verwenden, können Sie auf einen  _asyncResult_-Ausgabeparameter zugreifen und die  **status**- und  **error**-Eigenschaften des  **AsyncResult**-Parameterobjekts verwenden, um den Status des asynchronen Abrufs zu prüfen sowie um zu prüfen, ob der Abruf Fehlermeldungen hervorgebracht hat.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## So werden Empfänger hinzugefügt


Wenn Sie bereits bestehende Empfänger in einem Termin oder einer Nachricht nicht überschreiben möchten, können Sie anstelle von  **Recipients.setAsync** die asynchrone Methode **Recipients.addAsync** verwenden, um Empfänger hinzuzufügen. **addAsync** funktioniert insofern ähnlich wie **setAsync**, als dass es ein  _recipients_-Eingabeargument benötigt. Alternativ können Sie eine Rückrufmethode bereitstellen sowie jedes Argument für die Rückrufmethode, dass den Parameter "asyncContext" verwendet. Anschließend können Sie Status, Ergebnis und Fehlermeldungen des asynchronen  **addAsync**-Abrufs prüfen, indem Sie den  _asyncResult_-Ausgabeparameter der Rückrufmethode verwenden. Im folgenden Beispiel wird geprüft, ob das Element, das gerade verfasst wird, ein Termin ist, dem zwei Teilnehmer hinzugefügt werden.


```
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## Zusätzliche Ressourcen



- [Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook](60901b90-2472-4132-9b5e-0749e2e33cbe.md)
    
- [Abrufen und Festlegen von Outlook-Elementdaten in Formularen zum Lesen oder Verfassen](081ef8af-fc0b-4b1c-86e0-4ce8392e9886.md)
    
- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md)
    
- [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md)
    
- [Abrufen oder Festlegen des Betreffs, wenn Sie einen Termin oder eine Nachricht in Outlook verfassen](654b41f7-97c6-4abd-9eb4-3be337db82f2.md)
    
- [Einfügen von Daten in den Textkörper bei der Erstellung eines Termins oder einer Nachricht in Outlook](ca6d1571-70b4-4941-a10b-b0c730fcdf9a.md)
    
- [Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen](23558d28-e6f4-42e2-b943-27354c255a56.md)
    
- [Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen](7f3f8159-bf96-4dd4-aedc-e773e36904ea.md)
    
