
# Hinzufügen und Entfernen von Anhängen bei einem Element in einem Formular zum Verfassen in Outlook
Sehen Sie sich JavaScript-Beispiele für die Verwendung von asynchronen Methoden und einer Exchange-Webdienste-ID (EWS) aus einem Outlook-Add-In zum Hinzufügen einer Anlage zu einem Element an, das ein Benutzer in Outlook erstellt.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Datei oder Outlook-Element als Anlage hinzufügen


Zum Anhängen einer Datei oder eines Outlook-Elements an das Element, das gerade erstellt wird, können Sie die Methoden [addFileAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [addItemAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) verwenden. Beide Methoden sind asynchrone Methoden, das heißt, der Vorgang kann fortgesetzt werden, auch wenn die "add-attachment"-Aktion zum Hinzufügen der Anlage noch ausgeführt wird. Je nach dem ursprünglichen Speicherort und der Größe der hinzuzufügenden Anlage dauert die Ausführung des asynchronen "add-attachment"-Aufrufs einen Moment. Stehen Aufgaben an, die vom Abschluss der Aktion abhängig sind, müssen Sie diese Aufgaben mit einer Rückrufmethode ausführen. Diese Rückrufmethode ist optional und wird angestoßen, wenn der Upload der Anlage abgeschlossen ist. Die Rückrufmethode hat ein [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt als Ausgabeparameter, das Status, Fehler und Rückgabewert der "add-attachment"-Aktion liefert. Sollte der Rückruf Extraparameter erfordern, können Sie diese in einem optionalen  _options.aysncContext_-Parameter angeben.  _options.asyncContext_ kann von beliebiger Art sein, sofern die Rückrufmethode diese akzeptiert.

Beispiel: Sie legen  _options.asyncContext_ als JSON-Objekt fest, das mindestens ein Schlüsselwertpaar mit dem Trennzeichen ':' zwischen Schlüssel und Wert und dem Trennzeichen ',' zwischen den Schlüsselwertpaaren enthält. Weitere Beispiele zur [Übergabe von optionalen Parametern an asynchrone Methoden](../../docs/develop/asynchronous-programming-in-office-add-ins.md#AsyncProgramming_OptionalParameters) finden Sie auf der Office-Add-Ins Plattform in [Asynchrone Programmierung in Office-Add-Ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md). Im folgenden Beispiel sehen Sie, wie Sie den  **asyncContext**-Parameter zum Übergeben von zwei Argumenten an eine Rückrufmethode verwenden:




```
{ asyncContext: { var1: 1, var2: 2} }
```

Mit den Eigenschaften  **status** und **error** des **AsyncResult**-Objekts können prüfen, ob der asynchrone Methodenaufruf in der Rückrufmethode erfolgreich war oder nicht. War das Anhängen erfolgreich, verwenden Sie die  **AsyncResult.value**-Eigenschaft, um die Anlage-ID abzurufen. Die Anlage-ID ist eine Ganzzahl, die Sie später zum Entfernen der Anlage benötigen.


(../../images  Als bewährte Methode sollten Sie die Anlage-ID zum Entfernen von Anlagen nur dann verwenden, wenn dasselbe Add-In in derselben Sitzung diese Anlage angehängt hat. In Outlook Web App und OWA für mobile Geräte ist die Anlage-ID nur innerhalb derselben Sitzung gültig. Eine Sitzung gilt als beendet, wenn der Benutzer das Add-In schließt oder wenn der Benutzer in einem Inlineformular beginnt, ein Dokument zu verfassen und daher das Inlineformular aufruft, um in einem anderen Fenster seine Aktion fortzusetzen.


### Datei anhängen

Sie können einer Nachricht oder einem Termin in einer Entwurfsvorlage eine Anlage mit der  **addFileAttachmentAsync**-Methode hinzufügen und den Datei-URI dabei angeben. Ist die Datei geschützt, können Sie ein entsprechendes Identitäts- oder Authentifizierungstoken als Parameter der URI-Abfragezeichenfolge mit einschließen. Exchange ruft den URI auf, um die Anlage abzurufen, und der Webdienst, der die Datei schützt, verwendet das Token als Authentifizierung.

Das folgende JavaScript-Beispiel zeigt ein Entwurfs-Add-in, das eine Datei (picture.png) von einem Webserver an die Nachricht bzw. den Termin anhängt, die bzw. der gerade erstellt wird. Die Rückrufmethode verwendet  **asyncResult** als Parameter, prüft den Anhängestatus und ruft die Anlage-ID ab, falls das Anhängen erfolgreich ist.




```
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


### Outlook-Element anhängen

Sie können ein Outlook-Element (wie E-Mail, Kalender oder Kontakte) an eine Nachricht oder einen Termin in einer Entwurfsvorlage anhängen, indem Sie die EWS-ID (Exchange Web Services) des Elements angeben und die  **addItemAttachmentAsync**-Methode verwenden. Die EWS-ID eines E-Mail-, Kalender-, Kontakt- oder Aufgabenelements können Sie aus dem Benutzerpostfach mit der [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode abrufen, indem Sie auf den EWS-Vorgang [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx) zugreifen. Die [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft liefert Ihnen auch die EWS-ID eines vorhandenen Elements in einem Leseformular.

Die folgende JavaScript-Funktion,  `addItemAttachment`, erweitert das obige erste Beispiel und fügt ein Element als Anlage zu einer E-Mail oder einem Termin hinzu, die bzw. der gerade erstellt wird. Diese Funktion hat die EWS-ID des anzuhängenden Elements als Argument. Ist das Anhängen erfolgreich, ruft sie die Anlage-ID für eine weitere Verarbeitung in derselben Sitzung, wie Entfernen der Anlage, ab.




```
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


(../../images  Mit einem Entwurfs-Add-in können Sie eine Instanz eines periodischen Termins in Outlook Web App oder OWA für mobile Geräte anhängen. In einem unterstützenden Outlook-Rich-Client aber führt der Versuch, eine Instanz anzuhängen, zum Anhängen der Serie (Mastertermin).


## Anlage entfernen


Sie können eine angehängte Datei oder ein Element aus einer Nachricht oder einem Termin in einer Entwurfsvorlage entfernen, indem Sie die entsprechende Anlage-ID angeben und die Methode [removeAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) verwenden. Sie können nur Anlagen entfernen, die ein Add-In in ein und derselben Sitzung hinzugefügt hat. Vergewissern Sie sich, dass die Anlage-ID zu einer gültigen Anlage gehört, andernfalls würde die Methode einen Fehler zurückgeben. Ähnlich wie **addFileAttachmentAsync** und **addItemAttachmentAsync** ist **removeAttachmentAsync** eine asynchrone Methode. Mit einer Rückrufmethode und dem **AsyncResult**-Ausgabeparameterobjekt prüfen Sie Status und mögliche Fehler. Sie können über den optionalen  **asyncContext**-Parameter, der ein JSON-Objekt von Schlüsselwertpaaren ist, zusätzliche Parameter an die Rückrufmethode übergeben.

Die folgende JavaScript-Funktion,  `removeAttachment`, erweitert das obige Beispiel und entfernt die angegebene Anlage aus der E-Mail oder dem Termin, die bzw. der gerade erstellt wird. Diese Funktion hat die ID der zu entfernenden Anlage als Argument. Sie rufen die Anlage-ID nach einem erfolgreichen  **addFileAttachmentAsync**- oder  **addItemAttachmentAsync**-Methodenaufruf ab und speichern sie für einen späteren  **removeAttachmentAsync**-Methodenaufruf.




```
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## Tipps zum Hinzufügen und Entfernen von Anlagen


Wenn Ihr Entwurfs-Add-in Anlagen hinzufügt und entfernt, strukturieren Sie den Code derart, dass eine gültige Anlage-ID an den „remove-attachment"-Aufruf zum Entfernen der Anlage übergeben wird und Sie den Fall verarbeiten können, wenn  **AsyncResult.error** eine **InvalidAttachmentId** zurückgibt. Je nach dem Speicherort und der Größe der Anlage kann das Anhängen einer Datei oder eines Elements einige Zeit dauern. Im folgenden Beispiel ist ein Aufruf an **addFileAttachmentAsync**,  `write` und **removeAttachmentAsync** dargestellt. Sie gehen möglicherweise davon aus, dass die Aufrufe nacheinander abgearbeitet werden.


```
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

Auch wenn  **addFileAttachmentAsync** bevor **removeAttachmentAsync** startet, da **addFileAttachmentAsync**asynchron ist, können die Aufrufe  `write` und **removeAttachmentAsync** gestartet werden, bevor **addFileAttachmentAsync** beendet ist. Wenn dieser Fall eintritt, bleibt `attachmentID` **undefined**, und Sie erhalten eine Fehlermeldung für den  **removeAttachmentAsync**-Aufruf, wie in der folgenden Ausgabe gezeigt:




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

Um dies zu vermeiden, können Sie einerseits prüfen, ob  `attachmentID` definiert ist, bevor der **removeAttachmentAsync**-Aufruf erfolgt. Oder Sie initiieren den  **removeAttachmentAsync**-Aufruf aus der Rückrufmethode von  **addFileAttachmentAsync** heraus, wie es das folgende Beispiel darstellt:




```
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Ausgabebeispiel:




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

Beachten Sie, dass der Rückruf für  **removeAttachmentAsync** im Rückruf für **addFileAttachmentAsync** verschachtelt ist. Aufgrund der Tatsache, dass **addFileAttachmentAsync** und **removeAttachmentAsync** asynchron sind, kann die letzte Zeile im Rückruf für **addFileAttachmentAsync** ausgeführt werden, bevor der Rückruf für **removeAttachmentAsync** abgeschlossen ist.


## Zusätzliche Ressourcen



- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](../outlook/compose-scenario.md)
    
- [Asynchrone Programmierung in Office-Add-Ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


