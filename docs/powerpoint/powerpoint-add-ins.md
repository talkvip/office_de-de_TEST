
# Erstellen von Inhalts- und Aufgabenbereich-Add-Ins für PowerPoint
Entwickeln Sie Aufgabenbereich- und Inhalts-Add-Ins für PowerPoint.

 _**Gilt für:** apps for Office | Office Add-ins | PowerPoint_

Die Codebeispiele in diesem Artikel zeigen einige grundlegenden Aufgaben für die Entwicklung von PowerPoint-Inhalts-Add-Ins. Zum Anzeigen von Informationen hängen diese Beispiele von der  `app.showNotification`-Funktion ab, die in Visual StudioOffice-Add-Ins-Projektvorlagen enthalten ist. Wenn Sie Visual Studio nicht zum Entwickeln Ihres Add-Ins verwenden, müssen Sie die  `showNotification`-Funktion mit Ihrem eigenen Code ersetzen. Einige dieser Beispiele hängen ebenfalls von diesem  `globals`-Objekt ab, das außerhalb dese Bereichs folgender Funktionen deklariert wird:  `var globals = {activeViewHandler:0, firstSlideId:0};`

In diesen Codebeispielen ist es erforderlich, dass das Projekt auf [Office.js v1.1-Bibliothek oder neuer verweist ](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Erkennen der aktiven Ansicht der Präsentation und behandeln des ActiveViewChanged-Ereignisses

Die Funktion  `getFileView` ruft die Methode [Document.getActiveViewAsync ](http://msdn.microsoft.com/library/6b53c90a-df57-4851-98d1-fae2b54f6ad6%28Office.15%29.aspx) auf, um zurückzugeben, ob die aktuelle Ansicht der Präsentation „Bearbeiten" (alle Ansichten, in denen Sie Folien bearbeiten können, z. B. die Ansicht **Normal** oder die **Gliederungsansicht**) oder „Lesen" ist ( **Bildschirmpräsentation** oder **Leseansicht**)anzeigen.


```
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

Die Funktion  `registerActiveViewChanged` ruft die Methode [addHandlerAsync](http://msdn.microsoft.com/library/8b2ec6c4-0983-4f5e-abd9-16f15b4fc87b%28Office.15%29.aspx) auf, um einen Handlern für das Ereignis [Document.ActiveViewChanged](http://msdn.microsoft.com/library/f86afe63-bf70-43dd-b224-3bc53b5e991f%28Office.15%29.aspx) auf. Wenn Sie nach der Ausführung dieser Funktion die Ansicht der Präsentation ändern, zeigt die Benachrichtigung `app.showNotification` den aktiven Ansichtsmodus an („Lesen" oder „Bearbeiten").




```
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## Abrufen der URL der Präsentation

Die Funktion  `getFileUrl` ruft die Methode [Document.getFileProperties](http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4%28Office.15%29.aspx) auf, um die URL der Präsentationsdatei abzurufen.


```
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## Navigieren zu einer bestimmten Folie in der Präsentation

Die Funktion  `getSelectedRange` ruft die Methode [Document.getSelectedDataAsync](http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx) auf, um ein JSON-Objekt abzurufen, das von `asyncResult.value` zurückgegeben wurde, das ein Array mit der Bezeichnung "Folien" enthält, das wiederum die IDs, Titel und Indizes des ausgewählten Bereichs von Folien enthält. Sie speichert außerdem die ID der ersten Folie im ausgewählten Bereich in einer globalen Variable.


```
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

Die Funktion  `goToFirstSlide` ruft die Methode [Document.goToByIdAsync](http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380%28Office.15%29.aspx) auf, um die ID der ersten Folie abzurufen, die von der Funktion `getSelectedRange` oben gespeichert wurde.




```
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## Navigieren zwischen Folien in der Präsentation

Die  `goToSlideByIndex`-Funktion ruft die  **Document.goToByIdAsync**-Methode auf, um zur nächsten Folie in der Präsentation zu navigieren.


```
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## Weitere Aufgaben

Weitere Codebeispiele finden Sie in den folgenden Artikeln:


- [Lesen ausgewählter Daten](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md#ReadWriteDocumentData_Read)
    
- [Schreiben von Daten in die Auswahl](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md#ReadWriteDocumentData_Write)
    
- [Erkennen von Änderungen in der Auswahl](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md#ReadWriteDocumentData_DetectChanges)
    
- [So speichern Sie Add-in-Status und -Einstellungen für Inhalts- und Aufgabenbereich-Add-ins](persisting-add-in-state-and-settings.md#PersistSettingsContentTaskPaneApp)
    

## 



- [Lesen und Schreiben von Daten in die aktive Auswahl in einem Dokument oder Arbeitsblatt](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Abrufen des gesamten Dokuments aus einem Add-In für PowerPoint oder Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Verwenden von Dokumentdesigns in PowerPoint-Add-Ins](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
