
# Abrufen oder Festlegen des Betreffs, wenn Sie einen Termin oder eine Nachricht in Outlook verfassen
Erfahren Sie, wie Sie den Betreff aus einem Outlook-Add-In abrufen oder festlegen können, wenn der Benutzer einen Termin in Outlook erstellt.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Voraussetzungen zum Abrufen oder Festlegen des Betreffs in einer Entwurfsvorlage


Die JavaScript-API für Office stellt asynchrone Methoden ([subject.getAsync](../reference/outlook/Subject.html%28Office.15%29.md) und [subject.setAsync](https://dev.outlook.com/reference/add-ins/Subject.md%28Office.15%29.md)) zum Abrufen und Festlegen des Speicherorts eines Termins bereit, den der Benutzer erstellt. Diese asynchronen Methoden sind nur in Erstellungs-Add-Ins verfügbar. Vergewissern Sie sich, dass Sie das Add-In-Manifest entsprechend für Outlook eingerichtet haben, damit das Add-In in Erstellformularen aktiviert wird, siehe Abschnitt [Einrichten von Outlook-Add-Ins für Erstellformulare](e4126e58-4ddc-4891-9f19-aa6c1a258027.md#mod_off15_CreatingForCompose_SettingUp) von [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md).

Die Eigenschaft  **subject** ist für einen Lesezugriff sowohl in Erstell- als auch in Leseformularen für Termine und Nachrichten verfügbar. In einem Leseformular greifen Sie auf die Eigenschaft direkt aus dem übergeordneten Objekt zu, wie in:




```
item.subject
```

Da der Benutzer und Ihr Add-In in einem Erstellformular jedoch gleichzeitig den Betreff einfügen oder ändern können, müssen Sie die asynchrone Methode  **getAsync** verwenden, um den Betreff abzurufen, siehe unten:




```
item.subject.getAsync
```

Die Eigenschaft  **subject** ist für einen Schreibzugriff nur in Entwurfsvorlagen, und nicht in Leseformularen, verfügbar.

Wie die meisten asynchronen Methoden in der JavaScript-API für Office verwenden  **getAsync** und **setAsync** optionale Eingabeparameter. Weitere Informationen zum Festlegen dieser optionalen Eingabeparameter finden Sie unter [Übergeben optionaler Parameter an asynchrone Methoden](http://msdn.microsoft.com/de-de/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters) in [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md).


## Betreff abrufen


In diesem Abschnitt sehen Sie das Beispiel eines Codes, der den Betreff des/der vom Benutzer erstellten Termins/Nachricht abruft und den Betreff anzeigt. Bei diesem Codebeispiel wird von einer Regel im Add-In-Manifest ausgegangen, die das Add-In in einer Entwurfsvorlage für einen Termin/eine Nachricht aktiviert, wie unten dargestellt.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Wenn Sie  **item.subject.getAsync** verwenden möchten, stellen Sie eine Rückrufmethode bereit, die den Status und das Ergebnis des asynchronen Aufrufs prüft. Sie können beliebige Argumente für die Rückrufmethode über die optionalen _asyncContext_-Parameter bereitstellen. Sie rufen Status, Ergebnisse und etwaige Fehler über den  _asyncResult_-Ausgabeparameter des Rückrufs ab. Ist der asynchrone Aufruf erfolgreich, erhalten Sie über die Eigenschaft [AsyncResult.value](../reference/outlook/simple-types.md%28Office.15%29.md) den Betreff als einfache Textzeichenfolge.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Betreff festlegen


In diesem Abschnitt sehen Sie das Beispiel eines Codes, der den Betreff des/der vom Benutzer erstellten Termins/Nachricht festlegt. Ähnlich wie im vorherigen Beispiel geht dieses Codebeispiel von einer Regel im Add-In-Manifest aus, die das Add-In in einer Entwurfsvorlage für einen Termin/eine Nachricht aktiviert.

Wenn Sie  **item.subject.setAsync** verwenden möchten, geben Sie eine Zeichenfolge mit bis zu 255 Zeichen in den Datenparameter ein. Sie können auch eine Rückrufmethode und Argumente für die Rückrufmethode im _asyncContext_-Parameter bereitstellen. Prüfen Sie Status, Ergebnis und etwaige Fehlermeldungen im  _asyncResult_-Ausgabeparameter des Rückrufs. Ist der asynchrone Aufruf erfolgreich, fügt  **setAsync** die angegebene Betreffzeichenfolge als einfachen Text ein, der einen vorhandenen Betreff zu diesem Element überschreibt.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Zusätzliche Ressourcen



- [Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook](60901b90-2472-4132-9b5e-0749e2e33cbe.md)
    
- [Abrufen und Festlegen von Outlook-Elementdaten in Formularen zum Lesen oder Verfassen](081ef8af-fc0b-4b1c-86e0-4ce8392e9886.md)
    
- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md)
    
- [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md)
    
- [Abrufen, Festlegen oder Hinzufügen von Empfängern beim Erstellen eines Termins oder einer Nachricht in Outlook](c92f62f5-cd93-49c2-9c74-d9b195ee7ef4.md)
    
- [Einfügen von Daten in den Textkörper bei der Erstellung eines Termins oder einer Nachricht in Outlook](ca6d1571-70b4-4941-a10b-b0c730fcdf9a.md)
    
- [Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen](23558d28-e6f4-42e2-b943-27354c255a56.md)
    
- [Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen](7f3f8159-bf96-4dd4-aedc-e773e36904ea.md)
    
