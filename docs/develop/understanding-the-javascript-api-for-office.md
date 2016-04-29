
# Grundlegendes zur JavaScript-API für Office
Hier erfahren Sie Grundlegendes über die Funktionsbereiche der JavaScript-API für Office, die in der Datei „Office.js" implementiert ist.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Dieser Artikel enthält Informationen über die JavaScript-API für Office und ihre Verwendung. Referenzinformationen finden Sie unter [JavaScript API for Office](http://msdn.microsoft.com/de-de/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx). Informationen zum Ausführen und Bearbeiten von Code in Ihrem Webbrowser mit der JavaScript-API für Office in Excel Online finden Sie im [API Tutorial for Office](http://msdn.microsoft.com/en-us/office/dn449240.aspx). Weitere Informationen zum Aktualisieren von Visual Studio-Projektdateien auf die neueste Version der JavaScript-API für Office finden Sie unter [Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien](641dc473-0931-4e00-8164-e7808ceed64d.md).

 **Visuelles Untersuchen des Objektmodells der JavaScript-API für Office mit ZoomIt**

[![Hereinzoomen in JavaScript-API für Office](../../images/appIcons_all.png)](http://go.microsoft.com/fwlink/?LinkId=317268)
Zoom: [1.1](http://go.microsoft.com/fwlink/?LinkId=391751)
Untersuchen des Objektmodells nach Add-in-Typ oder Host: [1.1](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)

## Verweisen auf die Bibliothek der JavaScript-API für Office in Ihrer Add-in


Die Bibliothek der JavaScript-API für Office ist in die Datei Office.js und dazugehörige JS-Dateien implementiert, die anwendungsspezifische Implementierungen wie z. B. Excel-15.js und Outlook-15.js enthalten. Verweisen Sie innerhalb des  `<head>`-Tags der Website auf die Bibliothek der [JavaScript-API für Office](e76f608d-1900-4911-bd6f-aebe61510c96.md) (z. B. eine HTML-, ASPX- oder PHP-Datei), die die Benutzeroberfläche Ihres Add-ins mithilfe eines `script`-Tags implementiert. Das Attribut des Tags  `src` muss auf die folgende CDN-URL festgelegt werden:


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"/>
```

Dadurch wird ein Download der Dateien der JavaScript-API für Office beim ersten Laden Ihres Add-ins gestartet, um zu gewährleisten, dass die neueste Implementierung der Datei Office.js und die dazugehörigen Dateien für die angegebene Version verwendet werden.




## Initialisierung Ihres Add-ins


 **Betrifft:** Alle Add-in-Typen

Die JavaScript-API für Office bietet das [Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx)-Objekt, mit dem Entwickler einen Listener für das [Initialize](http://msdn.microsoft.com/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignis einer Office-Add-In implementieren können. Wenn die API für das Add-in bereit ist, mit der Interaktion mit Inhalten des Benutzers zu beginnen, löst sie das  **Office.initialize**-Ereignis aus. Sie können mithilfe von Code im  **initialize**-Ereignishandler gängige Add-in-Initialisierungsszenarien implementieren. Sie können beispielsweise den Benutzer auffordern, in Excel einige Zellen auszuwählen, und anschließend ein Diagramm einfügen, das mit den ausgewählten Werten initialisiert wird. Der initialize-Ereignishandler dient auch zum Initialisieren von anderen benutzerdefinierten Logiken in Ihrer Add-in, z. B. zum Einrichten von Bindungen, Anfordern der Eingabe von Standardwerten für Add-in-Einstellungen usw.

 **Wichtig:** Selbst wenn Ihr Add-in keine auszuführenden Initialisierungsaufgaben hat, müssen Sie mindestens eine minimale **Office.initialize** Ereignishandlerfunktion wie im folgenden Beispiel einbauen.




```
Office.initialize = function () {
};
```

Wenn Sie keinen  **Office.initialize**-Ereignishandler einbauen, meldet das Add-in möglicherweise einen Fehler beim Starten. Auch wenn ein Benutzer versucht, das Add-in mit einem Office Online-Webclient wie Excel Online, PowerPoint Online oder Outlook Web App zu verwenden, wird die Ausführung scheitern. 

Falls das Add-in mehr als eine Seite enthält, muss beim Laden einer neuen Seite diese Seite immer einen  **Office.initialize**-Ereignishandler enthalten oder aufrufen.

Weitere Details zur Ereignisabfolge beim Initialisieren eines Add-ins finden Sie unter [Laden des DOM und der Laufzeitumgebung](dfbf251c-c943-45c0-8305-f479a2d8cc7b.md).

Der  _reason_-Parameter der Listenerfunktion des  **initialize**-Ereignisses bietet Aufgabenbereichs- und Inhalts-Add-Ins (aber nicht Outlook-Add-Ins) Zugriff auf die [InitializationReason](http://msdn.microsoft.com/library/3a1fb60c-a6a7-4c73-b1d0-97096946382e%28Office.15%29.aspx)-Aufzählung, die angibt, wie die Initialisierung erfolgt ist. Ein Aufgabenbereichs- oder Inhalts-Add-in kann beispielsweise initialisiert werden, da der Benutzer es über das Menüband des Office-Clients eingefügt hat oder weil ein Dokument geöffnet wurde, das das Add-in bereits enthält.

Sie müssen mithilfe des Werts der  **InitializationReason**-Aufzählung eine unterschiedliche Logik dafür implementieren, ob das Add-in erstmals eingefügt wurde oder bereits im Dokument vorhanden ist. Das folgende Beispiel zeigt eine einfache Logik, die Sie dem vorherigen Beispiel hinzufügen können. Dabei wird der Wert des  _reason_-Arguments zum Anzeigen verwendet, wie das Aufgabenbereichs- oder Inhalts-Add-in initialisiert wurde.




```
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    // Display initialization reason.
    if (reason == "inserted")



    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## "Context"-Objekt


 **Betrifft:** Alle Add-in-Typen

Wenn ein Add-in initialisiert wird, verfügt es über viele verschiedene Objekte, mit denen es in der Laufzeitumgebung interagieren kann. Der Laufzeitkontext des Add-ins wird in der API vom [Context](http://msdn.microsoft.com/library/662883d5-b86f-4bdc-99f0-9ee9129ed16c%28Office.15%29.aspx)-Objekt angegeben. Das  **Context**-Objekt ist das Hauptobjekt, das Zugriff auf die wichtigsten Objekte der API bietet, z. B. die Objekte [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) und [Mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx), die wiederum Zugriff auf den Dokument- und Postfachinhalt bieten.

In Aufgabenbereichs- oder Inhalts-Add-ins können Sie über die [document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx)-Eigenschaft des  **Context**-Objekts auf die Eigenschaften und Methoden des  **Document**-Objekts zugreifen, um mit dem Inhalt von Word-Dokumenten, Excel-Arbeitsblättern oder Project-Zeitplänen zu interagieren. Gleichsam können Sie in Outlook-Add-ins über die [mailbox](http://msdn.microsoft.com/library/481fadd4-a5b2-46e7-9fad-b119f3c505dd%28Office.15%29.aspx)-Eigenschaft des  **Context**-Objekts auf die Eigenschaften und Methoden des  **Mailbox**-Objekts zugreifen, um mit dem Nachrichten-, Besprechungsanfragen- oder Termininhalt zu interagieren. 

Das  **Context**-Objekt bietet auch Zugriff auf die Eigenschaften [contentLanguage](http://msdn.microsoft.com/library/4fd063c2-0cd0-4b5b-8993-93d7ff8ce3bf%28Office.15%29.aspx) und [displayLanguage](http://msdn.microsoft.com/library/732ba34c-c99f-4c00-836d-4250eb7f0dac%28Office.15%29.aspx), mit denen Sie das Gebietsschema (die Sprache) des Dokuments oder Elements bzw. der Hostanwendung bestimmen können. Über die [roamingSettings](http://msdn.microsoft.com/library/fbe0c9ab-a281-46c0-a181-5b70c61a7153%28Office.15%29.aspx)-Eigenschaft können Sie auf die Elemente des [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx)-Objekts zugreifen.


## "Document"-Objekt


 **Betrifft:** Add-ins vom Typ „Inhalt" und „Aufgabenbereich"

Zum Interagieren mit Dokumentdaten in Excel, PowerPoint und Word bietet die API das [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)-Objekt. Über die Elemente des  **Document**-Objekts können Sie auf folgende Weisen auf Daten zugreifen: 


- Lesen und Schreiben in aktive Auswahlen in Form von Text, zusammenhängenden Zellen (Matrizen) oder Tabellen.
    
- Tabellendaten (Matrizen oder Tabellen).
    
- Bindungen (werden mit den "add"-Methoden des  **Bindings**-Objekts erstellt).
    
- Benutzerdefinierte XML-Komponenten (nur bei Word).
    
- Einstellungen oder Add-in-Status (pro Add-in im Dokument dauerhaft gespeichert).
    
Mithilfe des  **Document**-Objekts können Sie mit Daten in Project-Dokumenten interagieren. Die Project-spezifische Funktionalität der API ist in den Elementen der abstrakten Klasse [ProjectDocument](http://msdn.microsoft.com/library/1908af4f-93b9-4859-87e3-06942014fae1%28Office.15%29.aspx) dokumentiert. Weitere Informationen zum Erstellen von Aufgabenbereichs-Add-ins für Project finden Sie unter [Aufgabenbereich-Add-ins für Project](74e80cc5-8095-4d42-886b-47a0820e9e09.md).

Alle diese Formen des Datenzugriffs beginnen mit einer Instanz des abstrakten Objekts  **Document**.

Sie können auf eine Instanz des  **Document**-Objekts zugreifen, wenn das Aufgabenbereichs- oder Inhalts-Add-in initialisiert wird, indem Sie die [document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx)-Eigenschaft des  **Context**-Objekts verwenden. Das  **Document**-Objekt definiert allgemeine Datenzugriffsfunktionen, die von Word- und Excel-Dokumenten gemeinsam genutzt werden. Es bietet auch Zugriff auf das  **CustomXmlParts**-Objekt für Word-Dokumente.

Das  **Document**-Objekt bietet Entwicklern vier Möglichkeiten für den Zugriff auf Dokumentinhalte: 


- Auswahlbasierten Zugriff
    
- Bindungsbasierten Zugriff
    
- Benutzerdefinierten Zugriff auf Basis von XML-Elementen (nur Word)
    
- Zugriff auf Basis des gesamten Dokuments (nur PowerPoint und Word)
    
Um Ihnen zu verdeutlichen, wie die Methoden zum auswahl- und bindungsbasierten Zugriff funktionieren, möchten wir zunächst erklären, wie die Datenzugriffs-APIs bei den verschiedenen Office-Anwendungen einen einheitlichen Datenzugriff ermöglichen.


### Einheitlicher Datenzugriff in allen Office-Anwendungen

 **Betrifft:** Add-ins vom Typ „Inhalt" und „Aufgabenbereich"

Zum Erstellen von Erweiterungen, die in den verschiedenen Office-Dokumenten einheitlich funktionieren, abstrahiert die JavaScript-API für Office die Besonderheiten der einzelnen Office-Anwendungen mithilfe allgemeiner Datentypen und der Fähigkeit, für unterschiedliche Dokumentinhalte drei allgemeine Datentypen zu erzwingen.


#### Allgemeine Datentypen

Sowohl beim auswahl- als auch beim bindungsbasierten Datenzugriff werden Dokumentinhalte über Datentypen zur Verfügung gestellt, die alle unterstützten Office-Anwendungen gemeinsam haben. Von Office 2013 werden drei Hauptdatentypen unterstützt:



|**Datentyp**|**Beschreibung**|**Hostanwendungsunterstützung**|
|:-----|:-----|:-----|
|Text|Stellt eine Zeichenfolgendarstellung der Daten in der Auswahl bereit.|In Excel 2013, Project 2013 und PowerPoint 2013 wird nur unformatierter Text unterstützt. In Word 2013 werden drei Formate unterstützt: Nur-Text, HTML und Office Open XML (OOXML).Wenn Text in einer Zelle in Excel ausgewählt wird, lesen und schreiben auswahlbasierte Methoden den gesamten Inhalt der Zelle, auch wenn nur ein Teil des Texts in der Zelle ausgewählt ist. Wenn Text in Word ausgewählt ist, lesen und schreiben auswahlbasierte Methoden nur die Zeichenfolge, die ausgewählt ist.Project 2013 und PowerPoint 2013 unterstützen nur auswahlbasierten Datenzugriff.|
|Matrix|Stellt die Daten in der Auswahl oder Bindung als zweidimensionales  **Array**-Objekt bereit, das in JavaScript als ein Array von Arrays implementiert wird.Zwei Zeilen mit  **string**-Werten in zwei Spalten sind beispielsweise  ` [['a', 'b'], ['c', 'd']]`, und eine einzelne Spalte mit drei Zeilen ist beispielsweise  `[['a'], ['b'], ['c']]`.|Matrix-Datenzugriff wird nur in Excel 2013 und Word 2013 unterstützt.|
|Table|Stellt die Daten in der Auswahl oder Bindung als [TableData](http://msdn.microsoft.com/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx)-Objekt bereit. Das  **TableData**-Objekt stellt die Daten über die Eigenschaften  **headers** und **rows** zur Verfügung.|Tabellendatenzugriff wird nur in Excel 2013 und Word 2013 unterstützt.|

#### Datentyperzwingung

Die Datenzugriffsmethoden für die Objekte  **Document** und [Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) unterstützen die Angabe des gewünschten Datentyps mithilfe des _coercionType_-Parameters dieser Methoden und den dazugehörigen [CoercionType](http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx)-Enumerationswerte. Unabhängig von der tatsächlichen Form der Bindung unterstützen die verschiedenen Office-Anwendungen die allgemeinen Datentypen, indem die Umwandlung der Daten in den angeforderten Datentyp erzwungen wird. Wenn beispielsweise in Word eine Tabelle oder ein Absatz ausgewählt wird, kann der Entwickler angeben, dass diese/r als Nur-Text, HTML, Office Open XML oder Tabelle gelesen werden soll. Dabei übernimmt die API-Implementierung die erforderlichen Transformationen und Datenkonvertierungen. 


 >**Tipp**   **Wann sollten Sie die Matrix-Koersion der Tabellenkoersion für den Datenzugriff vorziehen?** Sofern Ihre tabulierten Daten beim Hinzufügen von Zeilen und Spalten dynamisch wachsen und Sie mit Tabellenüberschriften arbeiten müssen, sollten Sie den Tabellendatentyp (durch das Angeben des _coercionType_ Parameters eines **Document** oder **Binding**-Objektdatenzugriffsmethode als  `"table"` oder **Office.CoercionType.Table**). Das Hinzufügen von Zeilen und Spalten innerhalb der Datenstruktur wird sowohl für Tabellen- als auch Matrixdaten unterstützt, das Anfügen von Zeilen und Spalten wird jedoch nur für Tabellendaten unterstützt. Falls Sie nicht planen, Zeilen oder Spalten hinzuzufügen und für Ihre Daten keine Kopfzeilenfunktion erforderlich ist, sollten Sie den Matrix-Datentyp verwenden (durch das Angeben des  _coercionType_-Parameters der Datenzugriffsmethode als  `"matrix"` oder **Office.CoercionType.Matrix**), das ein einfacheres Modell der Interaktion mit den Daten bereitstellt.

Wenn die Daten nicht in den angegebenen Typ umgewandelt werden können, gibt die [AsyncResult.status](http://msdn.microsoft.com/library/eec9c712-79eb-4365-88a1-6d77649727c1%28Office.15%29.aspx)-Eigenschaft im Rückruf  `"failed"` zurück. Sie können dann über die [AsyncResult.error](http://msdn.microsoft.com/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx)-Eigenschaft auf ein [Error](http://msdn.microsoft.com/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx)-Objekt mit Informationen zu den Gründen des Misserfolgs des Methodenaufrufs zugreifen.


## Arbeiten mit Auswahlen beim "Document"-Objekt


Das  **Document** -Objekt stellt Methoden zur Verfügung, über die Sie die aktuelle Auswahl des Benutzers nach dem Motto "Abrufen und vergessen" lesen und in diese schreiben können. Hierzu bietet das **Document** -Objekt die Methoden **getSelectedDataAsync** und **setSelectedDataAsync**.

In den folgenden Themen finden Sie Codebeispiele, die das Anwenden von Aufgaben auf Auswahlen veranschaulichen:


## Arbeiten mit Bindungen mithilfe der Objekte "Bindings" und "Binding"


Der bindungsbasierte Datenzugriff ermöglicht Inhalts- und Aufgabenbereichs-Add-ins einen einheitlichen Zugriff auf einen bestimmten Bereich eines Dokuments oder Arbeitsblatts über eine einer Bindung zugeordnete ID. Das Add-in muss zunächst die Bindung herstellen, indem eine der Methoden aufgerufen wird, über die ein Teil des Dokuments einer eindeutigen ID zugeordnet wird: [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx) oder [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx). Nachdem die Bindung eingerichtet wurde, kann das Add-in über die bereitgestellte ID auf die Daten im zugeordneten Bereich des Dokuments oder Arbeitsblatts zugreifen. Das Erstellen von Bindungen bietet für Ihr Add-in folgenden Nutzen:


- Zugriff auf allgemeine Datenstrukturen in allen unterstützten Office-Anwendungen, z. B. Tabellen, Bereiche oder Text (eine fortlaufende Folge von Zeichen).
    
- Lese-/Schreibvorgänge, ohne dass der Benutzer eine Auswahl treffen muss.
    
- Herstellung einer Beziehung zwischen dem Add-in und den Daten im Dokument. Bindungen bleiben im Dokument erhalten, und es kann zu einem späteren Zeitpunkt auf diese zugegriffen werden.
    
Die Herstellung einer Bindung ermöglicht auch das Abonnieren von Daten- und Auswahländerungsereignissen in diesem bestimmten Bereich des Dokuments oder der Tabelle. Dies bedeutet, dass das Add-in nur von Änderungen benachrichtigt wird, die in dem gebundenen Bereich stattfinden und nicht von solchen, innerhalb des gesamten Dokuments oder der gesamten Tabelle.

Das [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)-Objekt macht eine [getAllAsync](http://msdn.microsoft.com/library/ef902b73-cc4c-4551-95de-d8a51eeba82f%28Office.15%29.aspx)-Methode verfügbar, die den Zugriff auf alle in dem Dokument oder in der Tabelle hergestellten Bindungen ermöglicht. Auf die einzelnen Bindungen kann über ihre ID unter Verwendung der [Bindings.getBindingByIdAsync](http://msdn.microsoft.com/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx)- oder der [Office.select](http://msdn.microsoft.com/library/23aeb136-da1f-4127-a798-99dc27bc4dae%28Office.15%29.aspx)-Methode zugegriffen werden. Mithilfe einer der folgenden Methoden des  **Bindings**-Objekts können neue Bindungen hergestellt und vorhandene entfernt werden: [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx), [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx) oder [releaseByIdAsync](http://msdn.microsoft.com/library/ad285984-8b44-435d-9b84-f0ade570c896%28Office.15%29.aspx).

Es gibt drei unterschiedliche Arten von Bindungen, die Sie mit dem  _bindingType_-Parameter angeben, wenn Sie eine Bindung mithilfe der Methode  **addFromSelectionAsync**, **addFromPromptAsync** oder **addFromNamedItemAsync** erstellen:



|**Bindungstyp**|**Beschreibung**|**Hostanwendungsunterstützung**|
|:-----|:-----|:-----|
|Text binding|Bindet einen Bereich des Dokuments, der als Text dargestellt werden kann.|In Word sind die meisten zusammenhängenden Auswahlen gültig, während in Excel nur Auswahlen mit einer Zelle das Ziel einer Textbindung sein können. In Excel wird nur unformatierter Text unterstützt. Word unterstützt drei Formate: Nur-Text, HTML und Open XML für Office.|
|Matrix binding|Bindet einen festen Bereich eines Dokuments, der tabulierte Daten ohne Kopfzeilen enthält.Daten in einer Matrixbindung werden als zweidimensionales  **Array**-Objekt geschrieben bzw. gelesen, das in JavaScript als ein Array von Arrays implementiert wird. Beispielsweise können zwei Zeilen mit  **string**-Werten in zwei Spalten als  ` [['a', 'b'], ['c', 'd']]` gelesen oder geschrieben werden, und eine einzelne Spalte mit drei Zeilen kann als `[['a'], ['b'], ['c']]` gelesen oder geschrieben werden.|In Excel kann eine beliebige Auswahl benachbarter Zellen für die Herstellung einer Matrixbindung verwendet werden. In Word wird die Matrixbindung nur für Tabellen unterstützt.|
|Table binding|Bindet einen Bereich eines Dokuments, der eine Tabelle mit Kopfzeilen enthält.Daten in einer Tabellenbindung werden als [TableData](http://msdn.microsoft.com/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx)-Objekt geschrieben oder gelesen. Das  **TableData**-Objekt stellt die Daten über die Eigenschaften  **headers** und **rows** zur Verfügung.|Jede Excel- oder Word-Tabelle kann die Basis einer Tabellenbindung sein. Nach Herstellen einer Tabellenbindung wird jede neue Zeile oder Spalte, die Benutzer der Tabelle hinzugefügt haben, automatisch in die Bindung einbezogen.|
Nachdem eine Bindung mithilfe einer der drei "add"-Methoden des  **Bindings** -Objekts erstellt wurde, können Sie mit den Daten und Eigenschaften der Bindung mithilfe der Methoden des entsprechenden Objekts arbeiten: [MatrixBinding](http://msdn.microsoft.com/library/35e8568e-9129-4c00-b30f-d8c3b2555f1e%28Office.15%29.aspx), [TableBinding](http://msdn.microsoft.com/library/1508795b-1c70-456c-b3bf-666d40cf8f50%28Office.15%29.aspx) oder [TextBinding](http://msdn.microsoft.com/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx). Diese drei Objekte erben die Methoden [getDataAsync](http://msdn.microsoft.com/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) und [setDataAsync](http://msdn.microsoft.com/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx) des **Binding**-Objekts, die Ihnen die Interaktion mit gebundenen Daten ermöglichen.

In den folgenden Themen finden Sie Codebeispiele, die veranschaulichen, wie Aufgaben mit Bindungen ausgeführt werden:


## Arbeiten mit benutzerdefinierten XML-Komponenten mithilfe der Objekte "CustomXmlParts" und "CustomXmlPart"


 **Betrifft:** Add-ins für den Aufgabenbereich

Die Objekte [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx) und [CustomXmlPart](http://msdn.microsoft.com/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx) der API bieten Zugriff auf benutzerdefinierte XML-Komponenten in Word, die die XML-gesteuerte Bearbeitung des Dokumentinhalts ermöglichen. Vorgehensweisen für das Arbeiten mit **CustomXmlParts** und **CustomXmlPart**-Objekten finden Sie im Codebeispiel [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## Arbeiten mit dem gesamten Dokument mithilfe der getFileAsync-Methode


 **Betrifft:** Add-ins für den Aufgabenbereich für Word und PowerPoint

Die [Document.getFileAsync](http://msdn.microsoft.com/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx)-Methode und Elemente der Objekte [File](http://msdn.microsoft.com/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx) und [Slice](http://msdn.microsoft.com/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx) sollen Funktionalität zum Aufteilen ganzer Word- und PowerPoint-Dokumentdateien in Segmente (Blöcke) von jeweils bis zu 4 MB bereitstellen. Weitere Informationen finden Sie unter [Vorgehensweise: Abrufen des gesamten Dateiinhalts aus einem Dokument in einem Add-in](47a4ab14-0f1e-4cc8-8814-fa7e97362360.md).


## "Mailbox"-Objekt


 **Betrifft:** Outlook-Add-ins

Outlook-Add-ins verwenden hauptsächlich eine Untergruppe der API, die über das [Mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx)-Objekt zur Verfügung gestellt wird. Für den Zugriff auf die Objekte und Elemente, die insbesondere in Outlook-Add-ins verwendet werden, z. B. das [Item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx)-Objekt, greifen Sie mithilfe der [mailbox](http://msdn.microsoft.com/library/481fadd4-a5b2-46e7-9fad-b119f3c505dd%28Office.15%29.aspx)-Eigenschaft des  **Context**-Objekts auf das  **Mailbox**-Objekt zu (siehe die folgende Codezeile).




```
// Access the Item object.
var item = Office.context.mailbox.item;

```

Darüber hinaus können Outlook-Add-ins die folgenden Objekte verwenden:


-  **Office**-Objekt (zur Initialisierung).
    
-  **Context**-Objekt (für Zugriff auf die Inhalts- und Anzeigespracheneigenschaften).
    
-  **RoamingSettings**-Objekt (zum Speichern von Outlook-Add-in-spezifischen benutzerdefinierten Einstellungen für das Postfach des Benutzers, in dem das Add-in installiert ist).
    
Informationen zum Verwenden von JavaScript in Outlook-Add-ins finden Sie unter [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md) und [Übersicht über Architektur und Features von Outlook-Add-Ins](2cd5641b-492b-4431-8388-7fc589163e9c.md).


## Matrix der API-Unterstützung


Diese Tabelle bietet eine Übersicht über die API und von verschiedenen Add-in-Typen unterstützten Features (Inhalt, Aufgabenbereich und Outlook) sowie über die Office-Anwendungen, die diese hosten können, wenn Sie [Office-Hostanwendungen, die von Ihrem Add-in unterstützt werden](cff9fbdf-a530-4f6e-91ca-81bcacd90dcd.md) mithilfe des [1.1-Add-in-Manifestschemas und Features angeben, die von der Version 1.1 JavaScript API für Office](641dc473-0931-4e00-8164-e7808ceed64d.md) unterstützt werden.


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**Hostname**|Datenbank|Arbeitsmappe|Postfach|Präsentation|Dokument|Projekt|
||**Unterstützte** **Hostanwendungen**|Access Web Apps|ExcelExcel Online|OutlookOutlook Web AppOWA für mobile Geräte|PowerPointPowerPoint Online|Word|Project|
|**Unterstützte Add-in-Typen**|Inhalt|v|v||v|||
||Aufgabenbereich||v||v|v|v|
||Outlook|||v||||
|**Unterstützte API-Features**|Text lesen/schreiben||v||v|v|vSchreibgeschützt|
||Matrix lesen/schreiben||v|||v||
||Tabelle lesen/schreiben||v|||v||
||HTML lesen/schreiben|||||v||
||Office Open XMLlesen/schreiben|||||v||
||Eigenschaften von Aufgaben, Ressourcen,Ansichten und Feldern lesen||||||v|
||SelectionChanged-Ereignisse||v|||v||
||Das gesamte Dokument abrufen||||v|v||
||Bindungenund Bindungsereignisse|vNur vollständige und teilweiseTabellenbindungen|v|||v||
||Benutzerdefinierte XML-Komponentenlesen/schreiben|||||v||
||Add-in-Statusdaten beibehalten(Einstellungen)|vPro Host-Add-in|vPro Dokument|vPro Postfach|vPro Dokument|vPro Dokument||
||SettingsChanged-Ereignisse|v|v||v|v||
||Aktiven Ansichtsmodus abrufenund geänderte Ereignisse anzeigen||||v|||
||Zu Speicherortenim Dokument navigieren||v||v|v||
||Kontextbezogen mithilfe vonRegeln und regulären Ausdrücken aktivieren|||v||||
||Elementeigenschaften lesen|||v||||
||Benutzerprofil lesen|||v||||
||Anlagen abrufen|||v||||
||Benutzeridentitätstoken abrufen|||v||||
||Exchange-Webdienste aufrufen|||v||||
