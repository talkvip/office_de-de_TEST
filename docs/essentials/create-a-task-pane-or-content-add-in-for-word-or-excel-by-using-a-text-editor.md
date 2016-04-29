
# Erstellen des ersten Aufgabenbereich- oder Inhalts-Add-ins für Word oder Excel mit einem Text-Editor
Erstellen Sie eine einfache Office-Add-In mithilfe eines Text-Editors erstellen.

 _**Gilt für:** apps for Office | Excel | Office Add-ins | Word_

Die einfachste Office-Add-In für Excel 2013 oder Word 2013 besteht aus einer XML-Manifestdatei, die auf eine Webseite oder Website verweist. In diesem Artikel wird beschrieben, wie Sie mit einem Text-Editor ein einfaches Add-in erstellen, das nur aus einer XML-Manifestdatei und einer HTML-Datei besteht.

Zum Implementieren dieses Add-Ins müssen Sie Folgendes erstellen:


- Eine HTML-Datei, mit der die Benutzeroberfläche des Add-ins implementiert wird.
    
- Eine XML-Manifestdatei, mit der die erforderlichen Metadaten zum Anzeigen und Ausführen des Add-ins in Word oder Excel definiert werden.
    
- Eine CSS-Datei zum Definieren eines Stylesheets für das Add-in.
    
- Eine Datei project.js mit JavaScript-Programmierlogik, die mithilfe der JavaScript-API für Office (Office.js) Datenzugriffsvorgänge für den Inhalt im Word- oder Excel-Dokument ausführen kann.
    

## Erstellen der Office-Add-In "Hello World"


Die Benutzeroberfläche des Add-ins wird von einer HTML-Datei bereitgestellt, die optional JavaScript-Programmierlogik enthalten kann. Diese ersten Schritte veranschaulichen, wie Sie das Add-in „Hello World" ohne Programmierlogik erstellen. Nachdem Sie das Add-in „Hello World" erstellt haben, erfahren Sie, wie Sie Programmierlogik für die Interaktion mit dem Inhalt des Dokuments oder der Arbeitsmappe hinzufügen.


### So erstellen Sie die Dateien für das Add-in „Hello World"


1. Erstellen Sie auf dem lokalen Laufwerk den Ordner HelloWorld (z. B. C:\HelloWorld). Speichern Sie alle in den folgenden Schritten erstellten Dateien in diesem Ordner.
    
2. Erstellen Sie die Datei HelloWorld.html, die den folgenden HTML-Code enthält.
    
  ```HTML
  <!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <link rel="stylesheet" type="text/css" href="program.css" />
    </head>
    <body>
        <p>Hello World!</p>
    </body>
</html>

  ```


    Diese Datei enthält den Mindestsatz an HTML-Tags zum Anzeigen der Benutzeroberfläche eines Add-ins.
    
3. Erstellen Sie die Datei program.css, die den folgenden CSS-Code enthält.
    
  ```
  body
{
    position:relative;
}
li :hover
{  
    text-decoration: underline;
    cursor:pointer;
}
h1,h3,h4,p,a,li
{
    font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    text-decoration-color:#4ec724;
}

  ```


    Mit dieser Datei wird das Stylesheet für das Add-in bereitgestellt.
    
4. Erstellen Sie die XML-Datei HelloWorld.xml, die den folgenden XML-Code enthält.
    
     >**Wichtig**  Ersetzen Sie den Wert im  `<id>`-Tag durch eine GUID, die Sie selbst generiert haben.

  ```XML
  <?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xsi:type="TaskPaneApp">
  <Id>08afd7fe-1631-42f4-84f1-5ba51e242f98</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Hello World add-in"/>
  <Description DefaultValue="My first app."/>
  <IconUrl DefaultValue=
    "http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>

  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="\\MyShare\MyManifests\HelloWorld\HelloWorld.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>

  ```


    Mit dieser Datei wird die XML-Manifestdatei für das Add-in bereitgestellt.
    
In den nächsten beiden Verfahren wird beschrieben, wie Sie Ihre Dateien in eine Netzwerkfreigabe kopieren und anschließend diesen Speicherort als vertrauenswürdigen Add-in-Katalog angeben, damit Sie Ihr Add-in testen können.


### So geben Sie einen vertrauenswürdigen Speicherort für das Manifest an


1. Erstellen Sie einen Ordner in einer Netzwerkfreigabe (z. B. \\MyShare\MyManifests). Achten Sie darauf, dass das  `<SourceLocation>`-Element der Manifestdatei HelloWorld.xml auf diesen Speicherort für die HTML-Seite des Add-ins verweist. Speichern Sie alle erstellten Dateien in diesem Ordner.
    
     >**Hinweis**  Alternativ können Sie nur die Manifestdatei HelloWorld.xml in dieser Freigabe speichern und dann die HTML-Datei auf einen Webserver kopieren. In diesem Fall müssen Sie sicherstellen, dass das  `<SourceLocation>`-Element der Manifestdatei HelloWorld.xml auf die URL der Datei HelloWorld.html auf diesem Server verweist.
2. Öffnen Sie ein neues Dokument in Excel oder Word.
    
3. Klicken Sie auf die Registerkarte  **Datei**, und klicken Sie dann auf  **Optionen**.
    
4. Wählen Sie  **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
    
5. Klicken Sie auf  **Vertrauenswürdige Add-in-Kataloge**.
    
6. Geben Sie in das Feld  **Katalog-URL** den Pfad der in Schritt 1 erstellten Netzwerkfreigabe ein, und wählen Sie dann **Katalog hinzufügen** aus.
    
7. Aktivieren Sie das Kontrollkästchen  **Im Menü anzeigen**, und klicken Sie dann auf  **OK**.
    
    Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird.
    
8. Schließen Sie Excel oder Word, und starten Sie Excel oder Word neu.
    

### So testen Sie das Add-in „Hello World" und führen es aus


1. Klicken Sie auf der Registerkarte  **Einfügen** auf **Meine Add-ins**. 
    
2. Klicken Sie im Dialogfeld  **Office-Add-ins** auf **Freigegebener Ordner**.
    
3. Wählen Sie  **Hello World-Add-in** und dann **Starten** aus.
    
    Das Add-in wird in einem Aufgabenbereich rechts des aktuellen Dokuments oder der aktuellen Arbeitsmappe geöffnet.
    

## Hinzufügen von Programmierlogik zum Add-in „Hello World"


In den nächsten Schritten erfahren Sie, wie Sie dem Add-in „Hello World" einfache Programmierlogik für die Interaktion mit dem Inhalt des Dokuments oder der Arbeitsmappe hinzufügen.


### So fügen Sie dem Add-in „Hello World" Programmierlogik hinzu


1. Öffnen Sie die Datei HelloWorld.html, und fügen Sie die  `<script>`-Tags innerhalb der  `<head>`-Tags der Datei hinzu.
    
  ```HTML
  <!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <link rel="stylesheet" type="text/css" href="program.css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script>
        <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
        <script src="Program.js"></script>
    </head>
    <body>
        <p>Hello World!</p>
    </body>
</html>
  ```


    Dadurch wird ein Verweis auf die Bibliotheksdatei Office.js hinzugefügt, mit der die JavaScript-API für Office implementiert wird. Darüber hinaus wird ein Verweis auf die Datei Program.js hinzugefügt, die wir für die Programmierlogik für das Add-in erstellen werden.
    
     >**Hinweis**  Das  `src`-Attribut des  `<script>`-Tags verweist auf die JavaScript-API für Office (office.js), die extern verfügbar sein wird. 
2. Ersetzen Sie  `<p>Hello World!</p>` durch die Zeilen innerhalb der `<body>`-Tags der Datei.
    
  ```HTML
  <!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <link rel="stylesheet" type="text/css" href="program.css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script>
        <script src="Program.js"></script>
    </head>
    <body>
        <button onclick="writeData()"> Write Data </button></br>
        <button onclick="ReadData()"> Read Selected Data </button></br>
        Results: <div id="results"></div>
    </body>
</html>

  ```


    Dadurch werden der Add-in-Benutzeroberfläche zwei Schaltflächen hinzugefügt und ein  `div`-Tag zum Anzeigen der Ergebnisse definiert.
    
3. Erstellen Sie die Datei Program.js, die den folgenden JavaScript-Code enthält.
    
     >**Hinweis**  [Office.initialize](http://msdn.microsoft.com/de-de/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx) muss als Funktion am Beginn der Codedatei definiert werden, damit die [Office.context](http://msdn.microsoft.com/de-de/library/6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf%28Office.15%29.aspx)-Eigenschaft verfügbar ist, wenn sie von den nachfolgenden Funktionen aufgerufen wird.

  ```
  // The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}
var MyArray = [['Berlin'],['Munich'],['Duisburg']];

function writeData() {
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: 'matrix' });
}

function ReadData() {
    Office.context.document.getSelectedDataAsync("matrix", function (result) {
        if (result.status === "succeeded"){
            printData(result.value);
        }

        else{
            printData(result.error.name + ":" + err.message);
        }
    });
}

      function printData(data) {
    {
        var printOut = "";

        for (var x = 0 ; x < data.length; x++) {
            for (var y = 0; y < data[x].length; y++) {
                printOut += data[x][y] + ",";
            }
        }
       document.getElementById("results").innerText = printOut;
    }
}

  ```

4. Stellen Sie die Add-in-Dateien wie im Verfahren „So geben Sie einen vertrauenswürdigen Speicherort für das Manifest an" beschrieben erneut bereit.
    
5. Fügen Sie die Add-In wie im Verfahren „So testen Sie das Add-in „Hello World" und führen es aus" beschrieben ein, und testen Sie sie.
    

## Weitere Ressourcen



- [Aufgabenbereich- und Inhalts-Add-Ins für Office 2013](task-pane-and-content-add-ins.md)
    
- [Designrichtlinien für Office-Add-Ins](../add-in-design.md)
    
- [Entwicklungslebenszyklus von Office-Add-Ins](../../docs/design/add-in-development-lifecycle.md)
    
- [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)
    
- [Grundlegendes zur JavaScript-API für Office](../develop/understanding-the-javascript-api-for-office.md)
    
- [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md)
    
- [API und Schemaverweise für Office-Add-Ins](../../reference/reference.md)
    
