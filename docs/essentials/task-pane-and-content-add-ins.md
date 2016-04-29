
# Aufgabenbereich- und Inhalts-Add-Ins für Office 2013
Verwenden Sie die JavaScript-API, um Aufgabenbereich- oder Inhalts-Add-ins mit den gewünschten Features für unterschiedliche Hostanwendungen zu erstellen. 

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | PowerPoint | Project | Word_


## Unterstützung der JavaScript-API für Inhalts- und Aufgabenbereich-Add-ins


In diesem Abschnitt wird kurz auf den Teil der [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx) eingegangen, den Sie in Inhalts- und Aufgabenbereich-Add-Ins aufrufen können. Unter [Grundlegendes zur JavaScript-API für Office](../develop/understanding-the-javascript-api-for-office.md) finden Sie eine Übersicht über die Features der API. Beispiele finden Sie unter [Codebeispiele für Office-Add-Ins](code-samples.md).

Verwenden Sie die folgende Tabelle zum Erkunden der API nach Add-in-Typ oder Hostanwendung



|**Erkunden nach Add-in-Typ**|**Erkunden nach Hostanwendung**|
|:-----|:-----|
|
|||
|:-----|:-----|
|[![Hereinzoomen in das Office-Objektmodell für Inhalts-Apps](../../images/appIcons_content.png)](http://go.microsoft.com/fwlink/?LinkId=391752)|Inhalts-Add-ins [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391752)|
|[![Vergrößern Sie das Objektmodell, um Apps im Aufgabenbereich anzuzeigen](../../images/appIcons_taskpane.png)](http://go.microsoft.com/fwlink/?LinkId=391757)|Aufgabenbereich-Add-ins [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391757)|
||[Laden Sie die Zuordnungen herunter](http://www.microsoft.com/en-us/download/details.aspx?id=42032)die für jeden Add-in-Typ und für jede Hostanwendung verfügbar sind.|
|
|||
|:-----|:-----|
|[![Access](../../images/appIcons_Access.png)](http://go.microsoft.com/fwlink/?LinkId=391750)|Access [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391750)|
|[![Hereinzoomen in das App-Objektmodell für Excel](../../images/appIcons_Excel.png)](http://go.microsoft.com/fwlink/?LinkId=391753)|Excel [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391753)|
|[![Hereinzoomen in das App-Objektmodell für PowerPoint](../../images/appIcons_PowerPoint.png)](http://go.microsoft.com/fwlink/?LinkId=391755)|PowerPoint [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391755)|
|[![Hereinzoomen in das App-Objektmodell für Project](../../images/appIcons_Project.png)](http://go.microsoft.com/fwlink/?LinkId=391756)|Project [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391756)|
|[![Hereinzoomen in das App-Objektmodell für Word](../../images/appIcons_Word.png)](http://go.microsoft.com/fwlink/?LinkId=391758)|Word [ZoomIt](http://go.microsoft.com/fwlink/?LinkId=391758)|
|
Sie können die von Inhalts- und Aufgabenbereich-Add-ins unterstützten primären Objekte und Methoden wie folgt kategorisieren:


1.  **Allgemeine Objekte, die gemeinsam mit anderen Office-Add-Ins** genutzt werden
    
    Zu diesen Objekten gehören [Office](http://msdn.microsoft.com/de-de/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx), [Context](http://msdn.microsoft.com/library/662883d5-b86f-4bdc-99f0-9ee9129ed16c%28Office.15%29.aspx) und [AsyncResult](http://msdn.microsoft.com/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx). Das  **Office**-Objekt ist das Stammobjekt der JavaScript-API für Office. Das  **Context**-Objekt stellt die Laufzeitumgebung der des Add-Ins dar. Sowohl  **Office** als auch **Context** dienen als grundlegende Objekte für jedes Office-Add-In. Das **AsyncResult**-Objekt stellt die Ergebnisse eines asynchronen Vorgangs dar, wie die Daten, die an die  **getSelectedDataAsync**-Methode zurückgegeben werden, die liest, was ein Benutzer in einem Dokument ausgewählt hat.
    
2.  **Das Objekt "Document"**
    
    Der Großteil der API-Teilmenge, die für Inhalts- und Aufgabenbereich-Add-ins verfügbar ist, wird durch die Methoden, Eigenschaften und Ereignisse des Objekts [Document](http://msdn.microsoft.com/de-de/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) verfügbar gemacht. Mithilfe dieser Teilmenge der API kann Ihre Inhalts- oder Aufgabenbereich-App die weiter unten in diesem Thema beschriebenen Aufgaben durchführen.
    
    Ein Inhalts- oder Aufgabenbereich-Add-In kann die [Office.context.document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx)-Eigenschaft verwenden, um auf das  **Document**-Objekt und über dieses auf die zentralen Elemente der API zum Arbeiten mit Daten in Dokumenten zuzugreifen, wie die Objekte [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) und [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx) und die Methoden [getSelectedDataAsync](http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx), [setSelectedDataAsync](http://msdn.microsoft.com/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx) und [getFileAsync](http://msdn.microsoft.com/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx). Das  **Document**-Objekt bietet außerdem die [mode](http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00%28Office.15%29.aspx)-Eigenschaft, um festzustellen, ob ein Dokument schreibgeschützt ist oder sich im Bearbeitungsmodus befindet, die [url](http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc%28Office.15%29.aspx)-Eigenschaft, um die URL des aktuellen Dokuments abzurufen, sowie Zugriff auf das [Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)-Objekt. Das  **Document**-Objekt unterstützt auch das Hinzufügen von Ereignishandlern für das [SelectionChanged](http://msdn.microsoft.com/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx)-Ereignis, sodass Sie erkennen können, wenn ein Benutzer seine Auswahl im Dokument ändert.
    
    Ein Inhalts- oder Aufgabenbereich-Add-In kann das  **Document**-Objekt erst aufrufen, nachdem das DOM und die Laufzeitumgebung geladen wurden, normalerweise im Ereignishandler für das [Office.initialize](http://msdn.microsoft.com/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignis. Informationen zum Ablauf der Ereignisse, wenn ein Add-In initialisiert wird, und dazu, wie Sie überprüfen, ob das DOM und die Laufzeitumgebung erfolgreich geladen wurden, finden Sie unter [Laden des DOM und der Laufzeitumgebung](../../docs/develop/loading-the-dom-and-runtime-environment.md).
    
3.  **Objekte zum Arbeiten mit bestimmten Features**
    
    Um mit bestimmten Features der API zu arbeiten, kann Ihr Inhalts- oder Aufgabenbereich-Add-in mit den folgenden Objekten und Methoden arbeiten:
    
      - Verwenden Sie die Methoden des [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)-Objekts, um Bindungen zu erstellen oder abzurufen, und arbeiten Sie dann mit deren Daten unter Verwendung der Methoden und Eigenschaften des [Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx)-Objekts.
    
  - Verwenden Sie die Objekte [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx), [CustomXmlPart](http://msdn.microsoft.com/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx) und verwandte Objekte, um benutzerdefinierte XML-Komponenten in Word-Dokumenten zu erstellen und zu bearbeiten.
    
  - Verwenden Sie die Objekte [File](http://msdn.microsoft.com/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx) und [Slice](http://msdn.microsoft.com/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx), um eine Kopie des gesamten Dokuments zu erstellen, es in Blöcke oder Segmente zu unterteilen und dann die Daten in diese Segmente einzulesen oder zu übertragen.
    
  - Verwenden Sie das [Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)-Objekt, um benutzerdefinierte Daten wie Benutzereinstellungen und Add-In-Status zu speichern.
    

 >**Wichtig**  Einige API-Elemente werden nicht in allen Office-Anwendungen unterstützt, die Inhalts- und Aufgabenbereich-Add-ins hosten können. Informationen dazu, wie Sie herausfinden, welche Elemente unterstützt werden, finden Sie in den folgenden Ressourcen:

Einen Überblick über die Unterstützung der JavaScript-API für Office, die in Office-Hostanwendungen verfügbar ist, finden Sie in der [Matrix der API-Unterstützung](understanding-the-javascript-api-for-office.md#APIOverview_APISupportMatrix) unter [Grundlegendes zur JavaScript-API für Office](../develop/understanding-the-javascript-api-for-office.md).


## Lesen und Schreiben in eine aktive Auswahl


Sie können die aktuelle Auswahl eines Benutzers in einem Dokument, einem Tabellenblatt oder einer Präsentation lesen oder in diese schreiben. Je nach Hostanwendung für Ihr Add-In können Sie den Typ der Datenstruktur zum Lesen oder Schreiben als Parameter in der [getSelectedDataAsync](http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx)- und der [setSelectedDataAsync](http://msdn.microsoft.com/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx)-Methode des [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)-Objekts angeben. Sie können beispielsweise jede Art von Daten (Text, HTML, Tabellendaten oder Office Open XML) für Word, Text- und Tabellendaten für Excel und Text für PowerPoint und Projekt angeben. Sie können auch Ereignishandler erstellen, um Änderungen an der Auswahl des Benutzers zu erkennen. Mit dem folgenden Beispiel werden Daten mithilfe der  **getSelectedDataAsync**-Methode aus der Auswahl als Text abgerufen.


```
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
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

Weitere Informationen und Beispiele finden Sie unter [Lesen und Schreiben von Daten in die aktive Auswahl in einem Dokument oder Arbeitsblatt](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Einrichten einer Bindung an eine Region in einem Dokument oder Tabellenblatt


Wie im vorherigen Abschnitt beschrieben, können Sie die  **getSelectedDataAsync**- und die  **setSelectedDataAsync**-Methode verwenden, um die  _aktuelle_ Auswahl des Benutzers in einem Dokument, einem Tabellenblatt oder einer Präsentation zu lesen oder in diese zu schreiben. Wenn Sie jedoch über mehrere Sitzungen des Add-ins hinweg auf dieselbe Region in einem Dokument zugreifen möchten, ohne dass der Benutzer eine Auswahl vornehmen muss, sollten Sie zuerst eine Bindung an diese Region einrichten. Sie können Daten- und Auswahländerungsereignisse für die gebundene Region auch abonnieren.

Sie können eine Bindung mithilfe der Methoden [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx), [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx) oder [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx) des [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)-Objekts hinzufügen. Diese Methoden geben eine ID zurück, die Sie für den Zugriff auf Daten in der Bindung verwenden können, oder mit der Sie deren Datenänderungs- oder Auswahlereignisse abonnieren können.

Im folgenden Beispiel wird dem aktuell ausgewählten Text in einem Dokument mithilfe der  **Bindings.addFromSelectionAsync**-Methode eine Bindung hinzugefügt. 




```
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Weitere Informationen und Beispiele finden Sie unter [Einrichten einer Bindung an Regionen in einem Dokument oder Arbeitsblatt](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## Abrufen vollständiger Dokumente


Wenn Ihr Aufgabenbereich-Add-in in PowerPoint oder Word ausgeführt wird, können Sie die [Document.getFileAsync](http://msdn.microsoft.com/de-de/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx)-, die [File.getSliceAsync](http://msdn.microsoft.com/de-de/library/5a8a5cc2-e883-42cd-92ab-d63e10c4c707%28Office.15%29.aspx)- und die [File.closeAsync](http://msdn.microsoft.com/de-de/library/1ad5cebf-6feb-43ff-8b19-97d91132ab2b%28Office.15%29.aspx)-Methode verwenden, um ein ganze Präsentation oder ein gesamtes Dokument abzurufen.

Wenn Sie  **Document.getFileAsync** aufrufen, erhalten Sie eine Kopie des Dokuments in einem [File](http://msdn.microsoft.com/de-de/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx)-Objekt. Das  **File**-Objekt bietet Zugriff auf das Dokument in "Segmenten", die als [Slice](http://msdn.microsoft.com/de-de/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx)-Objekte dargestellt werden. Wenn Sie  **getFileAsync** aufrufen, können Sie den Dateityp (Text oder komprimiertes Open Office XML-Format) und die Größe der Segmente (bis zu 4 MB) angeben. Um auf die Inhalte des **File**-Objekts zuzugreifen, können Sie dann die  **File.getSliceAsync**-Methode aufrufen, die die Rohdaten in der [Slice.data](http://msdn.microsoft.com/de-de/library/95a68949-6009-49ae-a531-2df77687b85d%28Office.15%29.aspx)-Eigenschaft zurückgibt. Wenn Sie das komprimierte Format angegeben haben, erhalten Sie die Dateidaten als Bytearray. Wenn Sie die Datei an einen Webdienst übertragen, können Sie die komprimierten Rohdaten vor der Übermittlung in eine Zeichenfolge mit Base64-Codierung umwandeln. Nachdem Sie die Segmente der Datei erhalten haben, können Sie das Dokument schließlich mithilfe der  **File.closeAsync**-Methode schließen.

Weitere Informationen finden Sie unter [Vorgehensweise: Abrufen des gesamten Dokuments aus einem Add-in für PowerPoint oder Word](http://msdn.microsoft.com/de-de/library/47a4ab14-0f1e-4cc8-8814-fa7e97362360%28Office.15%29.aspx). 


## Lesen und Schreiben benutzerdefinierter XML-Komponenten eines Word-Dokuments


Mithilfe des Open Office XML-Dateiformats und von Inhaltssteuerelementen können Sie benutzerdefinierte XML-Komponenten zu einem Word-Dokument hinzufügen und Elemente in den XML-Komponenten an Inhaltssteuerelemente in diesem Dokument binden. Wenn Sie das Dokument öffnen, übernimmt Word automatisch Daten aus den benutzerdefinierten XML-Komponenten in gebundene Inhaltssteuerelemente und liest diese. Benutzer können auch Daten in die Inhaltssteuerelemente schreiben. Wenn Benutzer das Dokument speichern, werden die Daten in den Steuerelementen auch in den gebundenen XML-Komponenten gespeichert. Aufgabenbereich-Add-ins für Word können die Eigenschaft [Document.customXmlParts](http://msdn.microsoft.com/de-de/library/b72c08bc-b49c-497c-9521-26ccce148bda%28Office.15%29.aspx) und die Objekte [CustomXmlParts ](http://msdn.microsoft.com/de-de/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx), [CustomXmlPart](http://msdn.microsoft.com/de-de/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx) und [CustomXmlNode](http://msdn.microsoft.com/de-de/library/dc1518de-47fa-4108-aab7-04a022724b04%28Office.15%29.aspx) verwenden, um Daten dynamisch in dem Dokument zu lesen und zu schreiben.

Benutzerdefinierte XML-Komponenten können Namespaces zugeordnet werden. Verwenden Sie zum Abrufen von Daten in einem Namespace die Methode [CustomXmlParts.getByNamespaceAsync](http://msdn.microsoft.com/de-de/library/9902f555-5c20-45d6-9a8c-ae6bf013dfaf%28Office.15%29.aspx). 

Sie können auch die Methode [CustomXmlParts.getByIdAsync](http://msdn.microsoft.com/de-de/library/31a21b58-426e-4bbe-acdf-885b32ce50ab%28Office.15%29.aspx) verwenden, um auf benutzerdefinierte XML-Komponenten anhand ihrer GUID zuzugreifen. Nachdem Sie eine benutzerdefinierte XML-Komponente abgerufen haben, verwenden Sie die Methode [CustomXmlPart.getXmlAsync](http://msdn.microsoft.com/de-de/library/6606365a-9244-49b5-9393-fe2186091af7%28Office.15%29.aspx) zum Abrufen der XML-Daten.

Um eine neue benutzerdefinierte XML-Komponente zu einem Dokument hinzuzufügen, rufen Sie mithilfe der Eigenschaft  **Document.customXmlParts** die benutzerdefinierten XML-Komponenten des Dokuments ab, und rufen Sie die Methode [CustomXmlParts.addAsync](http://msdn.microsoft.com/de-de/library/2816397c-b86a-4f52-8b13-036f527f4bb7%28Office.15%29.aspx) auf.

Ausführliche Informationen dazu, wie Sie mit einem Aufgabenbereich-Add-in mit benutzerdefinierten XML-Komponenten arbeiten, finden Sie unter [Erstellen besserer Add-ins für Word mit Office Open XML](http://msdn.microsoft.com/de-de/library/c5bad651-a42f-4e57-bc60-c9b27eb2383b%28Office.15%29.aspx).


## Beibehalten von Add-in-Einstellungen


Häufig müssen Sie benutzerdefinierte Daten für Ihr Add-in speichern, etwa Einstellungen eines Benutzers oder den Status eines Add-ins, und beim nächsten Öffnen des Add-ins auf diese Daten zugreifen. Zum Speichern dieser Daten können Sie gängige Techniken für die Webprogrammierung wie Browsercookies oder HTML 5-Webspeicher verwenden. Wenn Ihr Add-in in Excel, PowerPoint oder Word ausgeführt wird, können Sie alternativ die Methoden des [Document.Settings](http://msdn.microsoft.com/de-de/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)-Objekts verwenden. Mit dem  **Settings**-Objekt erstellte Daten werden in dem Tabellenblatt, der Präsentation oder dem Dokument gespeichert, in das das Add-in eingefügt wurde und mit dem sie gespeichert wurde. Diese Daten sind nur für das Add-in verfügbar, durch das sie erstellt wurden.

Um Roundtrips zum Server zu vermeiden, auf dem das Dokument gespeichert ist, werden mit dem  **Settings**-Objekt erstellte Daten zur Laufzeit im Speicher verwaltet. Zuvor gespeicherte Einstellungsdaten werden in den Speicher geladen, wenn das Add-in initialisiert wird, und Änderungen an den Daten werden nur wieder im Dokument gespeichert, wenn Sie die [Settings.saveAsync](http://msdn.microsoft.com/de-de/library/7147c221-937c-477c-98a6-f59d6200c27b%28Office.15%29.aspx)-Methode aufrufen. Die Daten werden intern in einem serialisierten JSON-Objekt als Namen-/Wertepaare gespeichert. Sie können die [get](http://msdn.microsoft.com/de-de/library/aeac06dd-994e-4235-b208-1bd117395296%28Office.15%29.aspx)-, [set](http://msdn.microsoft.com/de-de/library/4e2c9758-953e-41e8-aca6-d8daf764a584%28Office.15%29.aspx)- und [remove](http://msdn.microsoft.com/de-de/library/a92446bf-de65-45bd-8412-36ea8e77c5a2%28Office.15%29.aspx)-Methoden des  **Settings**-Objekts verwenden, um Elemente in der speicherinternen Kopie der Daten zu beschreiben, zu lesen oder zu löschen. Die folgende Codezeile zeigt, wie Sie eine Einstellung mit dem Namen  `themeColor` rstellen und deren Wert auf "grün" festlegen.




```
Office.context.document.settings.set('themeColor', 'green');
```

Da mit den Methoden  **set** und **remove** erstellte oder gelöschte Einstellungsdaten sich auf eine speicherinterne Kopie der Daten beziehen, müssen Sie **saveAsync** aufrufen, um Änderungen an Einstellungsdaten im Dokument, mit dem Ihr Add-in arbeitet, beizubehalten.

Weitere Informationen zum Arbeiten mit benutzerdefinierten Daten mithilfe der Methoden des  **Settings**-Objekts finden Sie unter [Beibehalten des Zustands und der Einstellungen von Add-ins](../../docs/develop/persisting-add-in-state-and-settings.md).


## Lesen der Eigenschaften eines Projektdokuments


Wenn Ihr Aufgabenbereich-Add-in in Project ausgeführt wird, kann Ihr Add-in Daten aus einigen der Projekt-, Ressourcen- und Aufgabenfelder im aktiven Projekt lesen. Dazu verwenden Sie die Methoden und Ereignisse des [ProjectDocument](http://msdn.microsoft.com/de-de/library/1908af4f-93b9-4859-87e3-06942014fae1%28Office.15%29.aspx)-Objekts, mit dem das  **Document**-Objekt erweitert wird, um zusätzliche Project-spezifische Funktionalität zur Verfügung zu stellen.

Beispiele für das Lesen von Project-Daten finden Sie unter [Erstellen des ersten Aufgabenbereich-Add-ins für Project 2013 mit einem Text-Editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Berechtigungsmodell und Steuerung


Ihr Add-in verwendet das  **Permissions**-Element in ihrem Manifest, um von der JavaScript-API für Office Zugriffsberechtigungen für die Funktionalitätsebene anzufordern, die sie benötigt. Wenn Ihre App z. B. Lese-/Schreibzugriff auf das Dokument benötigt, muss ihre App in ihrem  **Permissions**-Element  `ReadWriteDocument` als den Textwert angeben. Da Berechtigungen dazu da sind, einem Benutzer Datenschutz und Sicherheit zu bieten, sollten Sie als bewährte Vorgehensweise die Mindestberechtigungen anfordern, die sie für ihre Features benötigt. Das folgende Beispiel zeigt, wie Sie im Manifest einer Aufgabenbereich-App die Berechtigung **ReadDocument** anfordern.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
…<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
…
</OfficeApp>

```

Weitere Informationen finden Sie unter [Anfordern von Berechtigungen zur API-Verwendung in Inhalts- und Aufgabenbereich-Add-Ins](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## Weitere Ressourcen



- [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)
    
- [Schemareferenz für Office-Add-in-Manifeste (v1. 1)](http://msdn.microsoft.com/de-de/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [Problembehandlung von Benutzerfehlern mit Office Add-Ins](../../docs/testing/testing-and-troubleshooting.md)
    
