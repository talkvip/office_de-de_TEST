
# Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen
Erfahren Sie, wie Sie die Uhrzeit eines Termins aus einem Outlook-Add-In abrufen oder festlegen.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Voraussetzungen für das Abrufen oder Festlegen der Start- oder Endzeit in einem Formular zum Verfassen


Die JavaScript-API für Office stellt asynchrone Methoden ([Time.getAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md) und [Time.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)) zum Abrufen und Festlegen der Start- und Endzeit eines Termins bereit, den der Benutzer erstellt. Diese asynchronen Methoden sind nur in Erstellungs-Add-Ins verfügbar. Vergewissern Sie sich, dass Sie das Add-In-Manifest entsprechend für Outlook eingerichtet haben, damit das Add-In in Erstellformularen aktiviert wird, siehe Abschnitt [Einrichten von Outlook-Add-Ins für Erstellformulare](e4126e58-4ddc-4891-9f19-aa6c1a258027.md#mod_off15_CreatingForCompose_SettingUp) von [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md).

Die Eigenschaften [start](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) und [end](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) sind für Termine sowohl in Formularen zum Verfassen als auch zum Lesen verfügbar. In einem Leseformular können Sie die Eigenschaften direkt vom übergeordneten Objekt aus aufrufen, wie in:




```
item.start
```

und in:




```
item.end
```

In einem Formular zum Verfassen müssen Sie zum Abrufen der Start- oder Endzeit allerdings, wie unten gezeigt, die asynchrone Methode  **getAsync** verwenden, da der Benutzer und Ihr Add-In die Uhrzeit zur gleichen Zeit einfügen oder ändern können:




```
item.start.getAsync
```

und:




```
item.end.getAsync
```

Wie die meisten asynchronen Methoden in der JavaScript-API für Office verwenden  **getAsync** und **setAsync** optionale Eingabeparameter. Weitere Informationen zum Festlegen dieser optionalen Eingabeparameter finden Sie unter [Übergeben optionaler Parameter an asynchrone Methoden](http://msdn.microsoft.com/de-de/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters) in [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md).


## So rufen Sie die Start- oder Endzeit ab


Dieser Abschnitt enthält ein Codebeispiel, das die Startzeit des Termins abruft, den der Benutzer verfasst, und die Uhrzeit anzeigt. Sie können denselben Code verwenden und die Eigenschaft  **start** durch die Eigenschaft **end** ersetzen, um die Endzeit abzurufen. Bei diesem Codebeispiel wird davon ausgegangen, dass das Add-in-Manifest eine Regel enthält, welche das Add-In in einem Formular zum Verfassen für einen Termin aktiviert (siehe unten).


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Um  **item.start.getAsync** oder **item.end.getAsync** zu verwenden, stellen Sie eine Rückrufmethode bereit, die den Status und das Ergebnis des asynchronen Aufrufs überprüft. Sie können der Rückrufmethode über den optionalen Parameter _asyncContext_ jedes erforderliche Argument bereitstellen. Sie können den Status, Ergebnisse und Fehler mithilfe des Ausgabeparameters _asyncResult_ des Rückrufs abrufen. Wenn der asynchrone Aufruf erfolgreich ist, können Sie die Startzeit als **Date**-Objekt in UTC-Format mithilfe der [AsyncResult.value](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Eigenschaft abrufen.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## So legen Sie die Start- oder Endzeit fest


In diesem Abschnitt wird ein Codebeispiel gezeigt, bei dem die Startzeit des vom Benutzer verfassten Termins oder der Nachricht festgelegt wird. Sie können den gleichen Code verwenden und die  **start**-Eigenschaft durch die  **end**-Eigenschaft ersetzen, um die Endzeit festzulegen. Wenn das Formular zum Verfassen des Termins bereits eine Startzeit aufweist, wird durch das spätere Festlegen der Startzeit die Endzeit angepasst, sodass die vorherige Dauer des Termins bestehen bleibt. Wenn das Formular zum Verfassen des Termins bereits eine Endzeit aufweist, werden durch das spätere Festlegen der Endzeit die Dauer und die Endzeit angepasst. Wenn der Termin als ein ganztägiges Ereignis festgelegt wurde, wird die Endzeit durch das Festlegen der Startzeit auf 24 Stunden später festgelegt, und die UI für das ganztägige Ereignis im Formular zum Verfassen deaktiviert.

Ähnlich wie im vorherigen Beispiel wird in diesem Codebeispiel eine Regel in dem Add-In-Manifest vorausgesetzt, die das Add-In in einem Erstellformular eines Termins aktiviert.

Um  **item.start.setAsync** oder **item.end.setAsync** zu verwenden, geben Sie einen **Date**-Wert in UTC im  _dateTime_-Parameter an. Wenn Sie ein Datum auf Basis einer Eingabe des Benutzers auf dem Client erhalten, können Sie [mailbox.convertToUtcClientTime](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md) verwenden, um den Wert in ein **Date** -Objekt in UTC umzuwandeln. Sie können im _asyncContext_-Parameter eine optionale Rückrufmethode und Argumente für die Rückrufmethode angeben. Sie sollten den Status, Ergebnisse und Fehlermeldungen im  _asyncResult_-Ausgabeparameter des Rückrufs prüfen. Wenn der asynchrone Aufruf erfolgreich ist, fügt  **setAsync** die angegebene Start- und Endzeitzeichenfolge als Nur-Text ein und überschreibt dabei eine ggf. vorhandene Start- oder Endzeit für dieses Element.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
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
    
- [Abrufen oder Festlegen des Betreffs, wenn Sie einen Termin oder eine Nachricht in Outlook verfassen](654b41f7-97c6-4abd-9eb4-3be337db82f2.md)
    
- [Einfügen von Daten in den Textkörper bei der Erstellung eines Termins oder einer Nachricht in Outlook](ca6d1571-70b4-4941-a10b-b0c730fcdf9a.md)
    
- [Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen](23558d28-e6f4-42e2-b943-27354c255a56.md)
    
