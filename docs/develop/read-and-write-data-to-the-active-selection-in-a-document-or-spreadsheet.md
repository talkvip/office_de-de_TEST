
# Lesen und Schreiben von Daten in die aktive Auswahl in einem Dokument oder Arbeitsblatt
In diesem Artikel wird erläutert, wie Sie die Auswahl eines Benutzers in einem Dokument oder einer Arbeitsmappe lesen und in diese schreiben. Darüber hinaus wird beschrieben, wie Sie Ereignishandler für Änderungen an der Auswahl des Benutzers erstellen.

 _**Gilt für:** apps for Office | Excel | Office Add-ins | PowerPoint | Project | Word_


## Übersicht


Vom [Document](http://msdn.microsoft.com/de-de/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)-Objekt werden Methoden verfügbar gemacht, mit denen Sie die Auswahl eines Benutzers in einem Dokument oder einer Arbeitsmappe lesen und in diese schreiben können. Zu diesem Zweck weist das  **Document** -Objekt die Methoden **getSelectedDataAsync** und **setSelectedDataAsync** auf. In diesem Thema wird beschrieben, wie Sie Ereignishandler zum Erkennen von Änderungen an der Auswahl des Benutzers lesen, schreiben und erstellen.

Die Methode  **getSelectedDataAsync** gilt nur für die aktuelle Auswahl des Benutzers. Wenn die Auswahl im Dokument gespeichert werden soll, damit dieselbe Auswahl über Sitzungen des Add-ins hinweg zum Lesen und Schreiben verfügbar ist, müssen Sie mithilfe der Methode [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/de-de/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) eineBindung hinzufügen (oder mit einer der anderen „addFrom"-Methoden des Objekts [Bindings](http://msdn.microsoft.com/de-de/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) eine Bindung erstellen). Weitere Informationen dazu, wie Sie eine Bindung an einen Dokumentbereich erstellen und dann eine Bindung lesen und schreiben, finden Sie unter [Einrichten einer Bindung an Regionen in einem Dokument oder Arbeitsblatt](5bf788db-d788-4d91-bcb6-fc3913b40012.md).


### Lesen ausgewählter Daten


Im folgenden Beispiel wird gezeigt, wie Sie Daten aus einer Auswahl in einem Dokument mithilfe der [getSelectedDataAsync](http://msdn.microsoft.com/de-de/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx)-Methode abrufen.


```
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In diesem Beispiel ist der erste  _coercionType_-Parameter als  **Office.CoercionType.Text** angegeben (Sie können diesen Parameter auch mithilfe der Literalzeichenfolge `"text"`) angeben. Dies bedeutet, dass die [value](http://msdn.microsoft.com/de-de/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx)-Eigenschaft des [AsyncResult](http://msdn.microsoft.com/de-de/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx)-Objekts, das über den  _asyncResult_-Parameter in der Rückruffunktion verfügbar ist, ein  **string**-Objekt zurückgibt, das den im Dokument ausgewählten Text enthält. Durch Angabe anderer Koersionstypen ergeben sich unterschiedliche Werte. [Office.CoercionType](http://msdn.microsoft.com/de-de/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx) ist eine Enumeration verfügbarer Koersionstypwerte. **Office.CoercionType.Text** wird zur Zeichenfolge "text" ausgewertet.


 >**Tipp**   **Wann sollten Sie den Koersionstyp Matrix und wann den Koersionstyp Tabelle für den Datenzugriff verwenden?** Wenn die ausgewählten Tabellendaten dynamisch mit hinzugefügten Zeilen und Spalten wachsen sollen und Sie mit Tabellenkopfzeilen arbeiten müssen, sollten Sie den Datentyp Tabelle verwenden (indem Sie den _coercionType_-Parameter der  **getSelectedDataAsync**-Methode als  `"table"` oder **Office.CoercionType.Table** angeben). Das Hinzufügen von Zeilen und Spalten innerhalb der Datenstruktur wird in Tabellen- und Matrixdaten unterstützt, das Anfügen von Zeilen und Spalten jedoch nur für Tabellendaten.Wenn keine Zeilen und Spalten hinzugefügt werden sollen und keine Kopfzeilenfunktionen für die Daten erforderlich sind, sollten Sie den Datentyp Matrix verwenden (indem Sie den  _coercionType_-Parameter der  **getSelecteDataAsync**-Methode als  `"matrix"` oder **Office.CoercionType.Matrix** angeben), da dies eine einfachere Interaktion mit den Daten ermöglicht.

Die anonyme Funktion, die an die Funktion als zweiter  _callback_-Parameter übergeben wird, wird nach Abschluss des  **getSelectedDataAsync**-Vorgangs ausgeführt. Die Funktion wird mit einem einzelnen Parameter,  _asyncResult_, aufgerufen, der das Ergebnis und den Status des Aufrufs enthält. Falls der Aufruf fehlschlägt, ermöglicht die [error](http://msdn.microsoft.com/de-de/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx)-Eigenschaft des  **AsyncResult**-Objekts den Zugriff auf das [Error](http://msdn.microsoft.com/de-de/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx)-Objekt. Sie können den Wert der Eigenschaften [Error.name](http://msdn.microsoft.com/de-de/library/b76aaafd-bb34-4853-b29d-67adb1111b37%28Office.15%29.aspx) und [Error.message](http://msdn.microsoft.com/de-de/library/594e168e-4fdf-4e80-ba7e-4856a4a8ea5f%28Office.15%29.aspx) überprüfen, um festzustellen, weshalb der Vorgang fehlgeschlagen ist. Andernfalls wird der im Dokument ausgewählte Text angezeigt.

Die [AsyncResult.status](http://msdn.microsoft.com/de-de/library/eec9c712-79eb-4365-88a1-6d77649727c1%28Office.15%29.aspx)-Eigenschaft wird in der  **if**-Anweisung verwendet, um zu testen, ob der Aufruf erfolgreich war. [Office.AsyncResultStatus](http://msdn.microsoft.com/de-de/library/e2652105-03e8-4771-a985-e66c661fe3ea%28Office.15%29.aspx) ist eine Enumeration der verfügbaren **AsyncResult.status**-Eigenschaftswerte.  **Office.AsyncResultStatus.Failed** wird zur Zeichenfolge "failed" ausgewertet (und kann wiederum auch als diese Literalzeichenfolge angegeben werden).


### Schreiben von Daten in die Auswahl


Im folgenden Beispiel wird gezeigt, wie Sie für die Auswahl festlegen, dass "Hello World!" angezeigt wird.


```
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Wenn Sie unterschiedliche Objekttypen für den Parameter  _data_ übergeben, erhalten Sie auch unterschiedliche Ergebnisse. Das Ergebnis ist abhängig von der aktuellen Auswahl im Dokument, von welcher Anwendung das Add-in gehostet wird sowie ob die übergebenen Daten in die aktuelle Auswahl umgewandelt werden können.

Die anonyme Funktion, die an die [setSelectedDataAsync](http://msdn.microsoft.com/de-de/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx)-Methode als  _callback_-Parameter übergeben wird, wird nach Abschluss des asynchronen Aufrufs ausgeführt. Wenn Sie Daten mithilfe der  **setSelectedDataAsync**-Methode in die Auswahl schreiben, ermöglicht der  _asyncResult_-Parameter des Rückrufs nur den Zugriff auf den Status des Aufrufs und auf das [Error](http://msdn.microsoft.com/de-de/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx)-Objekt, falls der Aufruf fehlschlägt.

 **Hinweis:** Ab dieser Version von Excel 2013 SP1 und der entsprechenden Version von Excel Online können Sie nun [die Formatierung festlegen, wenn Sie eine Tabelle zur aktuellen Auswahl verfassen](46b05707-b350-41be-b6b8-311799c71a33.md).


### Erkennen von Änderungen in der Auswahl


Im folgenden Beispiel wird gezeigt, wie Sie Änderungen in der Auswahl mithilfe der [Document.addHandlerAsync](http://msdn.microsoft.com/de-de/library/8b2ec6c4-0983-4f5e-abd9-16f15b4fc87b%28Office.15%29.aspx)-Methode erkennen, um einen Ereignishandler für das [SelectionChanged](http://msdn.microsoft.com/de-de/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx)-Ereignis im Dokument hinzuzufügen.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Mit dem ersten  _eventType_-Parameter wird der Name des Ereignisses angegeben, das abonniert wird. Das Übergeben der Zeichenfolge  `"documentSelectionChanged"` für diesen Parameter entspricht dem Übergeben des **Office.EventType.DocumentSelectionChanged**-Ereignistyps der [Office.EventType](http://msdn.microsoft.com/de-de/library/82c79659-52da-48b0-92a9-831226eb9a7f%28Office.15%29.aspx)-Enumeration.

Die  `myHander()`-Funktion, die an die Funktion als zweiter  _handler_-Parameter übergeben wird, ist ein Ereignishandler, der ausgeführt wird, wenn die Auswahl im Dokument geändert wird. Die Funktion wird nach Abschluss des asynchronen Vorgangs mit einem einzelnen Parameter,  _eventArgs_, aufgerufen, der einen Verweis auf ein [DocumentSelectionChangedEventArgs](http://msdn.microsoft.com/de-de/library/283f2d97-2595-444b-86a6-286efd77f638%28Office.15%29.aspx)-Objekt enthält. Die [DocumentSelectionChangedEventArgs.document](http://msdn.microsoft.com/de-de/library/12974085-c146-45ff-aede-70e247d8426f%28Office.15%29.aspx)-Eigenschaft ermöglicht den Zugriff auf das Dokument, welches das Ereignis ausgelöst hat.


 >**Hinweis**  Sie können einem bestimmten Ereignis mehrere Ereignishandler anhängen, indem Sie erneut die  **addHandlerAsync**-Methode aufrufen und eine zusätzliche Ereignishandlerfunktion für den Parameter  _handler_ übergeben. Dies funktioniert ordnungsgemäß, solange der Name jeder Ereignishandlerfunktion eindeutig ist.


### Beenden der Erkennung von Änderungen in der Auswahl


Im folgenden Beispiel wird gezeigt, wie Sie die Überwachung des [Document.SelectionChanged](http://msdn.microsoft.com/de-de/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx)-Ereignisses beenden, indem Sie die [document.removeHandlerAsync](http://msdn.microsoft.com/de-de/library/47e0b00f-e301-4f21-836d-aeac783c42e0%28Office.15%29.aspx)-Methode aufrufen.


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Mit dem  `myHandler`-Funktionsnamen, der als zweiter  _handler_-Parameter übergeben wird, wird der Ereignishandler angegeben, der aus dem  **SelectionChanged**-Ereignis entfernt wird.


 >**Wichtig**  Wenn der optionale  _handler_-Parameter beim Aufrufen der  **removeHandlerAsync**-Methode weggelassen wird, werden alle Ereignishandler für das angegebene  _eventType_-Objekt entfernt.


## Weitere Ressourcen



- [Aufgabenbereich- und Inhalts-Add-Ins für Office 2013](baf73b23-4429-4b7f-bcb9-a99a9618ae38.md)
    
- [Lesen von Daten aus einer Bindung](5bf788db-d788-4d91-bcb6-fc3913b40012.md#BindRegions_Read)
    
- [Schreiben von Daten in eine Bindung](5bf788db-d788-4d91-bcb6-fc3913b40012.md#BindRegions_Write)
    
