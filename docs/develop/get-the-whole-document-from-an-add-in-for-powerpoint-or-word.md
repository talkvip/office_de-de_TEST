
# Abrufen des gesamten Dokuments aus einem Add-In für PowerPoint oder Word
Erstellen Sie ein Aufgabenbereich-Add-In für PowerPoint oder Word, das einen Verweis auf den gesamten Inhalt des aktuellen Dokuments abruft und an einen Webserver sendet.

 _**Gilt für:** apps for Office | Office Add-ins | PowerPoint | Word_

Sie können ein Office-Add-In erstellen, um das Senden oder Veröffentlichen eines Word 2013- oder PowerPoint 2013-Dokuments an einen Remotespeicherort per Mausklick zu ermöglichen. In diesem Artikel wird veranschaulicht, wie Sie ein einfaches Aufgabenbereich-Add-In für PowerPoint 2013 erstellen, mit dem die gesamte Präsentation als Datenobjekt abgerufen wird und diese Daten per HTTP-Anforderung an einen Webserver gesendet werden.

## Voraussetzungen für das Erstellen eines Add-ins für PowerPoint oder Word


In diesem Artikel wird angenommen, dass Sie zum Erstellen des Aufgabenbereich-Add-ins für PowerPoint oder Word einen Text-Editor verwenden. Dafür müssen Sie die folgenden Dateien erstellen:


- Auf einem freigegebenen Netzwerkordner bzw. auf einem Webserver benötigen Sie die folgenden Dateien:
    
      - Eine HTML-Datei (GetDoc_App.html), in der die Benutzeroberfläche und Links zu den JavaScript-Dateien (einschließlich "office.js" und host-spezifische .js-Dateien) und CSS-Dateien (Cascading Stylesheet) enthalten sind.
    
  - Eine JavaScript-Datei (GetDoc_App.js) für die Programmierlogik des Add-ins
    
  - Eine CSS-Datei (Program.css) für die Formatvorlagen und die Formatierung des Add-ins
    
- Eine XML-Manifestdatei (GetDoc_App.xml) für das Add-in, die auf einem freigegebenen Netzwerkordner oder in einem Add-in-Katalog verfügbar ist. Die Manifestdatei muss auf den Speicherort der oben erwähnten HTML-Datei verweisen.
    
Sie können ein Add-In für PowerPoint oder Word auch mithilfe von Visual Studio 2015 oder Napa Office 365-Entwicklungstools erstellen. Weitere Informationen zum Erstellen von Office-Add-Ins **REMOVE_ME** finden Sie in Tabelle 1.


### Wichtige Begriffe für die Erstellung eines Aufgabenbereich-Add-ins

Bevor Sie mit dem Erstellen dieses Add-Ins für PowerPoint oder Word beginnen, sollten Sie sich mit der Erstellung von Office-Add-Ins und der Verwendung von HTTP-Anforderungen vertraut gemacht haben. In diesem Artikel wird nicht beschrieben, wie Sie Base64-codierten Text aus einer HTTP-Anforderung auf einem Webserver decodieren. 




**Tabelle 1: Wichtige Begriffe für die Erstellung eines Aufgabenbereich-Add-ins für PowerPoint oder Word**


|**Titel des Artikels**|**Beschreibung**|
|:-----|:-----|
|[Erstellen eines Aufgabenbereich- oder Inhalts-Add-Ins mit Visual Studio](create-a-task-pane-or-content-add-in-with-visual-studio.md)|Beschreibt die Erstellung eines „Hello World"-Office-Add-Ins sowie dessen Erweiterung auf das Lesen und Schreiben sowie die Bindung an das Dokument.PowerPoint 2013Unterstützt nur die in diesem Artikel erwähnten Methoden  **getSelectedDataAsync** und **setSelectedDataAsync**.|
|[Erstellen des ersten Aufgabenbereich- oder Inhalts-Add-ins für Word oder Excel mit einem Text-Editor](create-a-task-pane-or-content-add-in-for-word-or-excel-by-using-a-text-editor.md)|Beschreibt die Erstellung eines einfachen Office-Add-Ins mit einem Text-Editor ohne weitere Tools.|
|[Erstellen eines Aufgabenbereich-Add-Ins mit Napa Office 365-Entwicklungstools](create-a-task-pane-add-in-with-napa.md)|Demonstriert die Erstellung eines einfachen Office-Add-Ins mithilfe der Napa Office 365-Entwicklungstools.|
|[XMLHttpRequest-Objekt](http://msdn.microsoft.com/library/ie/ms535874%28v=vs.85%29.aspx)|Beschreibt das  **XMLHttpRequest**-Objekt in JavaScript und enthält ein kurzes Codebeispiel.|
|[HttpRequest.InputStream-Eigenschaft](http://msdn.microsoft.com/library/system.web.httprequest.inputstream.aspx)|Beschreibt die Verwendung der  **HttpRequest.InputStream**-Eigenschaft zum Lesen des Texts einer HTTP-Anforderung auf einer ASP.NET-Webseite. Enthält auch ein kurzes Codebeispiel.|

## Erstellen des Manifests für das Add-in


Mit der XML-Manifestdatei für das Add-In für PowerPoint werden wichtige Informationen zum Add-In bereitgestellt: welche Anwendungen es hosten kann, der Speicherort der HTML-Datei, der Add-In-Titel und die Add-In-Beschreibung sowie viele weitere Merkmale.


1. Fügen Sie der Manifestdatei im Texteditor den folgenden Code hinzu.
    
  ```XML
  
<?xml version="1.0" encoding="utf-8" ?> 
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="TaskPaneApp">
    <Id>[Replace_With_Your_GUID]</Id> 
    <Version>1.0</Version> 
    <ProviderName>[Provider Name]</ProviderName> 
    <DefaultLocale>EN-US</DefaultLocale> 
    <DisplayName DefaultValue="Get Doc add-in" /> 
    <Description DefaultValue="My get PowerPoint or Word document add-in." /> 
    <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" /> 
    <Hosts>
      <Host Name="Document" /> 
      <Host Name="Presentation" /> 
    </Hosts>
    <DefaultSettings>
      <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" /> 
    </DefaultSettings>
    <Permissions>ReadWriteDocument</Permissions> 
</OfficeApp>
  ```

2. Speichern Sie die Datei unter dem Namen "GetDoc_App.xml" und mit UTF-8-Codierung an einem Netzwerkspeicherort oder in einem Add-in-Katalog.
    

## Erstellen der Benutzeroberfläche für das Add-in


Für die Benutzeroberfläche des Add-ins können Sie HTML-Daten verwenden, die direkt in die Datei "GetDoc_App.html" geschrieben werden. Die Programmierlogik und die Funktionen des Add-ins müssen in einer JavaScript-Datei (z. B. "GetDoc_Apps.js") enthalten sein.

Verwenden Sie das folgende Verfahren, um eine einfache Benutzeroberfläche für das Add-in zu erstellen, die eine Überschrift und eine einzelne Schaltfläche umfasst.


1. Fügen Sie in einer neuen Datei im Texteditor die folgenden HTML-Daten hinzu.
    
  ```HTML
  
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <title>Publish presentation</title>
        <link rel="stylesheet" type="text/css" href="Program.css" />
        <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>
        <script src="GetDoc_App.js"></script>
    </head>
    <body>
      <form>
        <h1>Publish presentation</h1>
        <br />
        <div><input id=’submit’ type="button" value="Submit" /></div>
        <br />
        <div><h2>Status</h2> 
            <div id="status"></div>
        </div>
      </form>
    </body>
</html>
  ```

2. Speichern Sie die Datei unter dem Namen "GetDoc_App.html" und mit UTF-8-Codierung an einem Netzwerkspeicherort oder auf einem Webserver.
    

 >**Hinweis**  Achten Sie darauf, dass die <head>-Tags des Add-ins gültige Links zur Datei "office.js" enthalten. 

Hier werden Cascading Stylesheets verwendet, um dem Add-in ein einfaches und gleichzeitig modernes und professionelles Aussehen zu verleihen. Verwenden Sie die folgenden CSS-Daten, um das Format des Add-ins zu definieren.


1. Fügen Sie in einer neuen Datei im Texteditor die folgenden CSS-Daten hinzu.
    
  ```
  
body
{
    font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
}
h1,h2
{
    text-decoration-color:#4ec724;
}
input [type="submit"], input [type="button"] 
{ 
    height:24px; 
    padding-left:1em; 
    padding-right:1em; 
    background-color:white; 
    border:1px solid grey; 
    border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0; 
    cursor:pointer; 
}
  ```

2. Speichern Sie die Datei unter dem Namen "Program.css" mit UTF-8-Codierung am Netzwerkspeicherort bzw. auf dem Webserver, auf dem sich die Datei "GetDoc_App.html" befindet.
    

## Hinzufügen des JavaScript-Codes zum Dokument


Im Code für das Add-in wird über einen Handler für das [Office.initialize](http://msdn.microsoft.com/de-de/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignis ein Handler für das click-Ereignis der Schaltfläche  **Submit** im Formular hinzugefügt, und der Benutzer wird informiert, dass das Add-in bereit ist.

Im folgenden Codebeispiel wird der Ereignishandler für das  **Office.initialize**-Ereignis zusammen mit der Hilfsfunktion  `updateStatus` zum Schreiben in das div-Element "status" veranschaulicht.




```
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

      // After the DOM is loaded, add-in-specific code can run.
      document.getElementById('submit').addEventListener("click",
          function () {
              sendFile();
          });}
      updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div. 
function updateStatus(message) {
    var statusInfo = document.getElementById("status");
    statusInfo.innerHTML += message + "<br/>";
}
```



Wenn Sie auf der Benutzeroberfläche die Schaltfläche  **Submit** betätigen, ruft das Add-in die `sendFile`-Funktion auf, in der ein Aufruf der Methode [Document.getFileAsync](http://msdn.microsoft.com/de-de/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx) enthalten ist. Für die Methode **getFileAsync** wird das asynchrone Muster auf ähnliche Weise wie bei anderen Methoden in der JavaScript-API für Office verwendet. Sie verfügt über den erforderlichen Parameter _fileType_ und die beiden optionalen Parameter _options_ und _callback_. 

Der  _fileType_-Parameter erwartet von der [FileType](http://msdn.microsoft.com/de-de/library/fadbb4cf-a0e4-47b2-93dd-123f0b06d4ae%28Office.15%29.aspx)-Enumeration drei Konstanten:  **Office.FileType.Compressed** ("compressed"), **Office.FileType.PDF** ("pdf") oder **Office.FileType.Text** ("text"). PowerPoint unterstützt nur **Compressed** als Argument, während Word alle drei Konstanten unterstützt. Wenn Sie **Compressed** für den _fileType_-Parameter übergeben, gibt die  **getFileAsync**-Methode das Dokument als PowerPoint 2013-Präsentationsdatei (*.pptx) oder Word 2013-Dokumentdatei (*.docx) zurück, indem auf dem lokalen Computer eine temporäre Kopie der Datei erstellt wird.

Die  **getFileAsync**-Methode gibt einen Verweis auf die Datei als [File](http://msdn.microsoft.com/de-de/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx)-Objekt zurück. Vom  **File**-Objekt werden vier Member verfügbar gemacht: die [size](http://msdn.microsoft.com/de-de/library/ac498911-a9b1-465a-8fc3-aba3735deb33%28Office.15%29.aspx)-Eigenschaft, die [sliceCount](http://msdn.microsoft.com/de-de/library/13171845-b077-432e-9c1b-f46f5c7ebeb8%28Office.15%29.aspx)-Eigenschaft, die [getSliceAsync](http://msdn.microsoft.com/de-de/library/5a8a5cc2-e883-42cd-92ab-d63e10c4c707%28Office.15%29.aspx)-Methode und die [closeAsync](http://msdn.microsoft.com/de-de/library/1ad5cebf-6feb-43ff-8b19-97d91132ab2b%28Office.15%29.aspx)-Methode. Die  **size**-Eigenschaft gibt die Byte-Anzahl der Datei zurück. Die  **sliceCount**-Eigenschaft gibt die Anzahl von [Slice](http://msdn.microsoft.com/de-de/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx)-Objekten (weitere Informationen hierzu unten) in der Datei zurück.

Verwenden Sie den folgenden Code, um das PowerPoint- oder Word-Dokument mithilfe der Methode  **Document.getFileAsync** als **File**-Objekt zurückzugeben und anschließend die lokal definierte Funktion  `getSlice` aufzurufen. Beachten Sie, dass das Objekt **File** (eine Zählervariable) und die Gesamtzahl an Segmenten in der Datei beim Aufrufen von `getSlice` mittels eines anonymen Objekts übergeben werden.




```

// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {

    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size +
                    " bytes");

                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
    });
}
```



Von der lokalen  `getSlice`-Funktion wird die  **File.getSliceAsync**-Methode aufgerufen, um ein Segment aus dem  **File**-Objekt abzurufen. Die  **getSliceAsync**-Methode gibt ein  **Slice**-Objekt aus der Auflistung der Segmente zurück. Es verfügt über die beiden erforderlichen Parameter  _sliceIndex_ und _callback_. Der  _sliceIndex_-Parameter nutzt eine ganze Zahl als Indexerstellung für die Auflistung der Segmente. Wie bei anderen Funktionen in der JavaScript-API für Office auch, verwendet die  **getSliceAsync**-Methode ebenfalls eine Rückruffunktion als Parameter für die Behandlung der Ergebnisse des Methodenaufrufs.

Über das  **Slice**-Objekt erlangen Sie Zugriff auf die Daten, die in der Datei enthalten sind. Falls im  _options_-Parameter der  **getFileAsync**-Methode nicht anders angegeben, hat das  **Slice**-Objekt eine Größe von 4 MB. Vom  **Slice**-Objekt werden drei Eigenschaften verfügbar gemacht: [size](http://msdn.microsoft.com/de-de/library/57fa1620-f0fb-415e-8e39-3d0f30feacf9%28Office.15%29.aspx), [data](http://msdn.microsoft.com/de-de/library/95a68949-6009-49ae-a531-2df77687b85d%28Office.15%29.aspx) und [index](http://msdn.microsoft.com/de-de/library/7a70fe31-27af-402f-960e-e0d47d344e83%28Office.15%29.aspx). Die  **size**-Eigenschaft ruft die Größe des Segments in Byte ab. Die  **index**-Eigenschaft ruft eine ganze Zahl ab, die für die Position des Segments in der Auflistung der Segmente steht. 




```

// Get a slice from the file and then call sendSlice.
function getSlice(state) {

    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {

            updateStatus("Sending piece " + (state.counter + 1) +
                " of " + state.sliceCount);

            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

Von der Eigenschaft  **Slice.data** werden die Rohdaten der Datei als Bytearray zurückgegeben. Wenn die Daten im Textformat (XML oder Nur-Text) vorliegen, enthält das Segment den Rohtext. Wenn Sie **Office.FileType.Compressed** für den Parameter _fileType_ von **Document.getFileAsync** übergeben, enthält das Segment die Binärdaten der Datei als Bytearray. Bei einer PowerPoint- oder Word-Datei enthalten die Segmente Bytearrays.

Sie müssen Ihre eigene Funktion implementieren (oder eine verfügbare Bibliothek verwenden), um Bytearraydaten in eine Base64-codierte Zeichenfolge zu konvertieren. Weitere Informationen zum Base64-Codieren mit JavaScript finden Sie unter [Base64-Codierung und Decodierung](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Nachdem Sie die Daten in das Base64-Format konvertiert haben, können Sie diese auf unterschiedliche Arten auf einen Webserver übertragen, z. B. auch als Text einer HTTP POST-Anforderung.

Fügen Sie den folgenden Code hinzu, um ein Segment an einen Webdienst zu senden.


 >**Hinweis**  Mit diesem Code wird eine PowerPoint- oder Word-Datei in mehreren Segmenten an den Webserver gesendet. Der Webserver bzw. Webdienst muss jedes einzelne Segment zu einer separaten PPTX-Datei kompilieren, bevor Sie Änderungen daran vornehmen können.




```

function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't 
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request 
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status 
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST 
        // request to the web server.
        request.send(fileData);
    }
}
```



Wie der Name vermuten lässt, wird die Methode  **File.closeAsync** zum Schließen der Verbindung zum Dokument und zum Bereitstellen von Ressourcen verwendet. Obwohl von der Sandbox des Office-Add-Inss für nicht mehr relevante Dateiverweise eine Speicherbereinigung durchgeführt wird, ist es trotzdem am besten, Dateien explizit zu schließen, nachdem die Bearbeitung durch Ihren Code beendet ist. Die Methode **closeAsync** verfügt über einen _callback_-Parameter, mit dem angegeben wird, welche Funktion nach Abschluss des Aufrufs aufgerufen wird.




```

function closeFile(state) {

    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```


## Zusätzliche Ressourcen



- [Aufgabenbereich- und Inhalts-Add-Ins für Office 2013](task-pane-and-content-add-ins.md)
    
- [File-Objekt](http://msdn.microsoft.com/de-de/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx)
    
- [Slice-Objekt](http://msdn.microsoft.com/de-de/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx)
    
- [Erstellen des ersten Aufgabenbereich- oder Inhalts-Add-ins für Word oder Excel mit einem Text-Editor](create-a-task-pane-or-content-add-in-for-word-or-excel-by-using-a-text-editor.md)
    
