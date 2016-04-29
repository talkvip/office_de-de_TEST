
# Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen
Erfahren Sie, wie Sie den Speicherort aus einem Outlook-Add-In abrufen oder festlegen können, wenn der Benutzer einen Termin in Outlook erstellt.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Voraussetzungen zum Abrufen oder Festlegen des Speicherorts in einem Erstellformular


Die JavaScript-API für Office stellt asynchrone Methoden ([getAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md) und [setAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)) zum Abrufen und Festlegen des Speicherorts eines Termins bereit, den der Benutzer erstellt. Diese asynchronen Methoden sind nur in Erstellungs-Add-Ins verfügbar. Vergewissern Sie sich, dass Sie das Add-In-Manifest entsprechend für Outlook eingerichtet haben, damit das Add-In in Erstellformularen aktiviert wird, siehe Abschnitt [Einrichten von Outlook-Add-Ins für Erstellformulare](e4126e58-4ddc-4891-9f19-aa6c1a258027.md#mod_off15_CreatingForCompose_SettingUp) von [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md).

Die [location](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft ist in Lese- und Erstellformularen von Terminen für den Lesezugriff verfügbar. In einem Leseformular haben Sie direkt über das übergeordnete Objekt Zugriff auf die Eigenschaft, wie beispielsweise in:




```
item.location
```

Da der Benutzer und Ihr Add-In in einem Erstellformular jedoch gleichzeitig den Speicherort einfügen oder ändern können, müssen Sie die asynchrone Methode  **getAsync** verwenden, um den Speicherort abzurufen, siehe unten:




```
item.location.getAsync
```

Die  **location**-Eigenschaft ist nur in Erstellformularen von Terminen für den Schreibzugriff, jedoch nicht in Leseformularen.

Wie die meisten asynchronen Methoden in der JavaScript-API für Office verwenden  **getAsync** und **setAsync** ptionale Eingabeparameter. Weitere Informationen zum Festlegen dieser optionalen Eingabeparameter finden Sie unter [Übergeben optionaler Parameter an asynchrone Methoden](http://msdn.microsoft.com/en-us/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters) in [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md).


## So rufen Sie den Speicherort ab


In diesem Abschnitt wird ein Codebeispiel vorgestellt, das den Speicherort eines vom Benutzer erstellten Termins abruft und anzeigt. In diesem Codebeispiel wird eine Regel in dem Add-In-Manifest vorausgesetzt, die das Add-In in einem Erstellformular eines Termins aktiviert, siehe unten.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Zum Verwenden von  **item.location.getAsync** müssen Sie eine Rückrufmethode bereitstellen, die den Status und das Ergebnis des asynchronen Aufrufs überprüft. Sie können der Rückrufmethode die erforderlichen Argumente über den optionalen Parameter _asyncContext_ bereitstellen. Mithilfe des Ausgabeparameters _asyncResult_ im Rückruf können Sie den Status, die Ergebnisse und eventuelle Fehler abrufen. Falls der asynchrone Aufruf erfolgreich erfolgt, können Sie den Speicherort mithilfe der [AsyncResult.value](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Eigenschaft als Zeichenfolge abrufen.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## So legen Sie den Speicherort fest


In diesem Abschnitt wird ein Codebeispiel vorgestellt, das den Speicherort des vom Benutzer erstellten Termins festlegt. Ähnlich wie im vorherigen Beispiel wird in diesem Codebeispiel eine Regel in dem Add-In-Manifest vorausgesetzt, die das Add-In in einem Erstellformular eines Termins aktiviert.

Geben Sie zum Verwenden von  **item.location.setAsync** eine Zeichenfolge von bis zu 255 Zeichen in dem Datenparameter an. Optional können Sie in dem _asyncContext_-Parameter eine Rückrufmethode und Argumente für die Rückrufmethode bereitstellen. Im  _asyncResult_-Ausgabeparameter des Rückrufs sollten Sie den Status, das Ergebnis und eventuelle Fehlermeldungen überprüfen. Falls der asynchrone Aufruf erfolgreich erfolgt, fügt  **setAsync** den angegebenen Speicherort als Nur-Text-Zeichenfolge ein und überschreibt damit den vorhandenen Speicherort für dieses Element.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
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
    
- [Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen](7f3f8159-bf96-4dd4-aedc-e773e36904ea.md)
    
