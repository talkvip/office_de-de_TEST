
# Formatieren von Tabellen in Add-Ins für Excel
Erfahren Sie, wie Sie Formatierungen für Tabellen in Add-ins für Excel angeben.

 _**Gilt für:** apps for Office | Excel | Office Add-ins_

In diesem Artikel werden die verschiedenen Funktionen der Formatierungs-API erläutert und Ihre Funktionsweise aufgezeigt. In dieser Version können Sie das Zellenformat und einige weitere Optionen nur für Tabellen (nicht für  **Office.CoercionType.Text**- oder  **Office.CoercionType.Matrix**-Datenstrukturen) und nur in Excel-Add-Ins angeben. So legen Sie die Formatierung in Ihrem Add-In fest:

- Der Benutzer wählt die Tabelle aus (oder die Stelle an der eine Tabelle programmgesteuert eingefügt werden soll), anschließend kann das Add-in die  **Document.setSelectedDataAsync**-Methode für die Tabelle aufrufen und die Formatierung festlegen. 
    
    Oder…
    
- Wenn die Arbeitsmappe bereits gebundene Tabellen enthält (oder Ihr Add-in eine der „addFrom"-Methoden des [Bindings](http://msdn.microsoft.com/de-de/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)-Objekt verwendet, um bei der Initialisierung gebundene Tabellen zu erstellen), kann Ihr Add-in die Methode  **Binding.setDataAsync** für die gebundenen Tabellen aufrufen, um die Formatierung festzulegen.
    
 **Wichtig:** Um diese neuen und aktualisierten Methoden für das Foramtieren von Tabellen in Excel-Add-Ins zu verwenden, muss Ihr Add-in-Projekt [Office.js v1.1 oder höher verwenden oder für die Verwendung davon aktualisiert werden](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

## Angeben von Formatierungen

Um das Format anzugeben, das Sie festlegen möchten, müssen Sie ein JavaScript-Objektliteral verwenden, das eine oder mehrere Schlüsselwertpaare enthält. Eine Reihe von Formatierungseinstellungen können in einer Liste kombiniert werden. Beispiel: 


```
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

Um die Formatierung anzuwenden, wird das JavaScript-Objekt anschließend an eine der Methoden weitergegeben, die Formatierungsdaten oder andere Funktionen in der Tabelle unterstützen.

Formatierungen können auf zwei Arten angewendet werden:


- Das Add-in schreibt erstmalig Daten in eine Auswahl oder Bindung, indem der optionale Parameter  _cellFormat_ oder _tableOptions_ im Objekt _options_ angegeben wird, das an die Methode [Document.setSelectedDataAysnc](http://msdn.microsoft.com/de-de/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx) oder [Binding.setDataAsync](http://msdn.microsoft.com/de-de/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx) übergeben wird.
    
- Nach dem Festlegen der Formatierung können Sie die Formatierung mit den neuen für diesen Zweck bestimmten Methoden [löschen oder aktualisieren](#aktualisieren-und-löschen-der-formatierung).
    

## Verwenden optionaler Parameter mit Methoden zum Festlegen von Daten

Bei Tabellenbindungen können Sie folgende optionale Parameter verwenden, um die Formatierung mit den  **Document.setSelectedData**- oder  **Binding.setDataAsync**-Methoden beim Einstellen der Daten anzugeben:  _tableOptions_ und _cellFormat_.


### Der optionale Parameter tableOptions

Verwenden Sie den optionalen  _tableOptions_-Parameter, um Standard-Tabellenformate anzuwenden und bestimmte Tabellenfunktionen zu aktivieren oder zu deaktivieren. Beispiel:  **Kopfzeile**,  **Ergebniszeile** und **verbundene Zeilen**. Der Wert, den Sie als  _tableOptions_-Parameter weitergeben, ist ein JavaScript-Objekt, das eine Liste von Schlüsselwertpaaren enthält, z. B. 


```
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### Der optionale Parameter cellFormat

Verwenden Sie den optionalen  _cellFormat_-Parameter zum Ändern von Zellenformatwerten wie z. B. die Breite, Höhe, Schriftart, Ausrichtung usw. Der Wert, den Sie als  _cellFormat_-Parameter weitergeben, ist ein Array, das eine Liste von JavaScript-Objekten enthält, die angeben, welche Zellen betroffen sind und welche Formate darauf angewendet werden sollen. Beispiel:


```
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

Sie können mehrere  `cells:`- und  `format:`-Paare in dem  _cellFormat_-Array kombinieren, um die Anzahl der zum Anwenden der Formatierung erforderlichen Funktionsaufrufe zu verringern. 


#### Zellen

Verwenden Sie  `cells:` zum Angeben des Bereichs der Spalten, Zeilen und Zellen, auf die Sie die Formatierung anwenden möchten.


**In Zellenwerten unterstützte Bereiche**


|**Zellenbereicheinstellungen**|**Beschreibung**|
|:-----|:-----|
| `{row: i}`|Gibt den Bereich an, der sich zu der. i-Zeile des Datenteils der Tabelle erweitert.|
| `{column: i}`|Gibt den Bereich an, der sich zu der. i-Spalte des Datenteils der Tabelle erweitert.|
| `{row: i, column: j}`|Gibt den Bereich der Zellen aus der. i-Zeile zur. j-Spalte des Datenteils der Tabelle an.|
| `Office.Table.All`|Gibt die gesamte Tabelle, einschließlich Spaltenüberschriften, Daten und Ergebnissen (falls vorhanden) an.|
| `Office.Table.Data`|Gibt nur die Daten in der Tabelle an (keine Überschriften und Ergebnisse).|
| `Office.Table.Headers`|Gibt nur die Kopfzeile an.|

#### Format

Verwenden Sie den Befehl  `format:`, um die Formatierung anzugeben, die Sie mit dem Befehl  `cells:` auf den als Liste von JavaScript-Schlüsselwertpaaren definierten Bereich anwenden möchten. Eine Liste der unterstützten Werte finden Sie unter [Unterstützte Formatierungsschlüssel und -werte](#unterstützte-formatierungsschlüssel-und--werte).

 **Grenzwerte für die Formatierung für Excel Online**

Wenn die Formatierung in Excel Online festgelegt wird, darf die Anzahl der  _Formatierungsgruppen_, die an den Parameter _cellFormat_ weitergegeben wird, 100 nicht überschreiten. Eine einzelne Formatierungsgruppe besteht aus einem Formatierungssatz, der auf einen bestimmten Zellenbereich angewendet wird. (In anderen Worten: alles, was in einer der `cells:`-Objektliterale im Array angegeben ist, was an  _cellFormat_ weitergegeben wird.) Folgender Aufruf gibt zum Beispiel zwei Formatierungsgruppen an _cellFormat_ weiter.




```
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### Anwenden von optionalen Parametern

In dieser Version unterstützen nur die  **Document.setSelectedDataAsync**- und  **TableBinding.setDataAsync**-Methoden das Schreiben von Daten und das Festlegen von Formatierungen von Tabellen in einem Aufruf mithilfe der optionalen  _tableOptions_- und  _cellFormat_-Parameter. In den folgenden Beispielen muss der an den ersten Parameter jeder einzelnen Methode weitergegebene Wert  `tableData` (der _data_-Parameter) ein [TableData](http://msdn.microsoft.com/de-de/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx)-Objekt sein, das die Definition der Tabelle und die zu schreibenden Daten enthält.

 **Document.setSelectedDataAsync-Beispiel**




```
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **TableBinding.setDataAsync-Beispiel**




```
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **Hinweis:** Der Aufruf an `Office.select("bindings#myBinding")` geht davon aus, dass in dem Arbeitsblatt bereits eine Bindung mit dem Namen `myBinding` vorhanden ist.


## Aktualisieren und Löschen der Formatierung


Beim Festlegen der Formatierung mit den optionalen  _cellFormat_- und  _tableOptions_-Parametern der  **Document.setSelectedDataAsync**- oder  **TableBinding.setDataAsync**-Methoden wird die Formatierung nur beim ersten Aufruf festgelegt. Zum Aktualisieren oder Löschen der Formatierung müssen Sie drei neue Methoden des  **TableBinding**-Objekts verwenden:  **setFormatsAsync**,  **setTableOptionsAsync** und **clearFormatsAsync**.


### Aktualisieren der Formatierungen

Die [TableBinding.setFormatsAsync](http://msdn.microsoft.com/de-de/library/49712906-f582-4055-9ef8-6edde6e97679%28Office.15%29.aspx)-Methode dient nur zum Aktualisieren der Zellenformatierungen, wie z. B. der Breite, Höhe, Schriftart, des Hintergrunds und der Ausrichtung. Sie verwendet  _cellFormat_ als erforderlichen Parameter:


```
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

Die [TableBinding.setTableOptionsAsync](http://msdn.microsoft.com/de-de/library/2885fc57-4527-4ca4-a43d-9ee447ec27d3%28Office.15%29.aspx)-Methode dient nur zum Aktualisieren der Tabellenoptionen, wie z. B. der verbundenen Zeilen und Filterschaltflächen. Sie verwendet  _tableOptions_ als erforderlichen Parameter:




```
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### Formatierung löschen

Die [TableBinding.clearFormatsAsync](http://msdn.microsoft.com/de-de/library/cc56e9c0-b33c-4d9b-b676-a7e50f757c10%28Office.15%29.aspx)-Methode dient zum Löschen aller Formatierungen in der Tabelle. Sie verwendet den optionalen  _asyncContext_-Parameter und eine optionale Rückruffunktion: 


```
Office.select("bindings#myBinding").clearFormatsAsync();
```


## Unterstützte Formatierungsschlüssel und -werte


In den folgenden Tabellen ist eine Liste der unterstützten Schlüsselwertpaare aufgeführt, die Sie an den Parameter  _cellFormat_ oder _tableOptions_ weitergeben können.

Für  `format:`-Werte entsprechen die verfügbaren Einstellungen einer Untergruppe der im Dialogfeld  **Zellen formatieren** (Rechtsklick auf > **Zellen formatieren** oder **Format** > **Zellen formatieren** auf der Registerkarte **Home** im Menüband). Für `tableOptions:`-Werte entsprechen die Einstellungen den Gruppen  **Optionen für Tabellenformat** und **Tabellenformatvorlagen** auf der Registerkarte **Tabellentools** |**Entwurf** im Menüband.


 >**Wichtig**  Die Methoden der Formatierungs-API unterstützen nur die unten zusammengefassten Optionen und Werte. Wenn Sie andere Formatierungsoptionen oder -werte angeben, ist das Verhalten der Behandlung undefiniert. Diese undefinierten Verhalten der Behandlung sind plattformübergreifend nicht unbedingt konsistent; Sie sollten Ihre Add-ins für eine beliebige Plattform nicht basierend auf etwaigen Nebeneffekten dieses undefinierten Verhaltens entwickeln. Die undefinierten Verhalten der Behandlung sollten jedoch keine negativen Auswirkungen auf den Status und die Benutzeroberfläche Ihres Add-ins oder die Dokumente, mit denen es interagiert, haben.


**Ausrichtung**


|**Schlüssel**|**Werte**|**Notizen**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"Standard" | "Links" | "Zentriert" | "Rechts" | "Ausfüllen" | "Blocksatz" | "Über Auswahl zentrieren" | "Verteilt"|Bei der Kombination mit einem Einzugswert werden nur folgende Kombinationen unterstützt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">alignHorizontal: "left"</span> und <span class="code">indentLeft: <value></span></p></li><li><p><span class="code">alignHorizontal: "right"</span> und <span class="code">indentRight: <value></span></p></li><li><p><span class="code">alignHorizontal: "distributed"</span> und <span class="code">indentDistributed: <value></span></p></li></ul>|
| `alignVertical:`|"Oben" | "Zentrieren" | "Unten" | "Blocksatz" | "Verteilt"||



**Hintergrund**


|**Schlüssel**|**Werte**|**Notizen**|
|:-----|:-----|:-----|
| `backgroundColor:`|"keine" | <Namen der vordefinierten Farben> | #RRGGBB|Namen der vordefinierten Farben:"Schwarz", "Blau", "Grau", "Grün", "Orange", "Rosa", "Lila", "Rot", "Blaugrün", "Türkis", "Violett", "Weiß", "Gelb"|



**Rahmen**


|**Schlüssel**|**Werte**|**Notizen**|
|:-----|:-----|:-----|
| `borderStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Vordefinierte Rahmenformatnamen:"Strich Punkt", "Strich Punkt Punkt", "Gestrichelt", "Punktiert", "Doppelt", "Haarstrich", "Strich Punkt Mittel", "Strich Punkt Punkt Mittel", "Gestrichelt Mittel", "Mittel", "Schrägstrich Strich Punkt", "Dick", "Dünn"Gilt für alle Rahmen im angegebenen Bereich. (Entspricht dem Angeben der Rahmenformatvorlagen mit den Voreinstellungen  **Außen** und **Innen** auf der Registerkarte **Rahmen** des Dialogfelds **Zellen formatieren**.) **Hinweis:** Excel 2013 unterstützt die Wiedergabe aller 13 vordefinierten Rahmenarten. Excel Online unterstützt jedoch nicht alle Rahmenarten. Die folgende Tabelle beschreibt die Wiedergabe, die für jede Rahmenart verwendet wird, wenn Sie die Arbeitsmappe in Excel Online öffnen.

|**Excel 2013**|**Excel Online**|
|:-----|:-----|
|"Strich-Punkt"|gestrichelt (1px)|
|"Strich-Punkt-Punkt"|gepunktet (1px)|
|"gestrichelt"|gepunktet (1px)|
|"gepunktet"|gestrichelt (1px)|
|"Double"|Double (3px)|
|"Hair"|Solid (1px)|
|"Strich Punkt Mittel"|Gestrichelt (2 Pixel)|
|"Strich Punkt Mittel"|gepunktet (2 Pixel)|
|"gestrichelt Mittel"|gestrichelt (2 Pixel)|
|"Medium"|Solid (2 Pixel)|
|"Schrägstrich-Strich-Punkt"|gestrichelt (2 Pixel)|
|"Dick"|Solid (3px)|
|"Dünn"|Solid (1px)|
|
| `borderColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderTopStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderTopColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderBottomStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderBottomColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderLeftStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt für alle Rahmen im angegebenen Bereich.|
| `bordeerLeftColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderRightStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderRightColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderOutlineStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderOutlineColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt für alle Rahmen im angegebenen Bereich.|
| `borderInlineStyle:`|"keine" | <Alle vordefinierten Rahmenformatnamen>|Gilt nur für innere Rahmenlinien im angegebenen Bereich. (Entspricht dem Angeben der Rahmenformatvorlagen nur mit der Voreinstellung  **Innen** auf der Registerkarte **Rahmen** des Dialogfelds **Zellen formatieren**.)|
| `borderInlineColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB|Gilt nur für die inneren Rahmenlinien im angegebenen Bereich |



**Breite, Höhe und Umbruch der Zellen**


|**Schlüssel**|**Werte**|
|:-----|:-----|
| `width:`|"Automatisch anpassen" |  **Number**|
| `height:`|"Automatisch anpassen" |  **Number**|
| `wrapping:`|**Boolean**|



**Schriftart**


|**Schlüssel**|**Werte**|**Notizen**|
|:-----|:-----|:-----|
| `fontFamily:`|<Alle verfügbaren Schriftartnamen>|Wenn Sie eine Schriftart in Excel Online festlegen und die Schriftart nicht im Browser verfügbar ist, wird der API versuchen, zurück zu den folgenden Schriftarten in dieser Reihenfolge zu fallen: Segoe UI, Thonburi, Arial, Verdana und Microsoft Sans Serif. Falls keine dieser Schriftarten verfügbar ist, wird die Standardschrift des Browsers verwendet.|
| `fontStyle:`|"Normal" | "Kursiv" | "Fett" | "Fett Kursiv"|
 >**Hinweis**  Zum Zeitpunkt dieser Veröffentlichung verhält sich das Einstellen von  `fontStyle:` zu "kursiv" und dann zu "fett" (oder umgekehrt) als Zusammenschluss dieser beiden Einstellungen. Das bedeutet in diesem Beispiel, dass aus der Einstellung "kursiv" und anschließend "fett" schließlich "fett kursiv" wird. Um einen Bereich _nur_ kursiv oder fett zu formatieren, der zuvor fett oder kursiv formatiert war, müssen Sie zunächst `fontStyle:"regular"` einstellen, um die vorherige Formatierung zu löschen.

|
| `fontSize:`|**Number**||
| `fontUnderlineStyle:`|"keine" | "Einfach" | "Doppelt" | "Einfach (Buchhaltung)" | "Doppelt (Buchhaltung)"||
| `fontColor:`|"automatisch" | <Alle Namen der vordefinierten Farben> | #RRGGBB||
| `fontDirection:`|"Kontext" | "Von links nach rechts" | "Von rechts nach links"|Excel Online unterstützt derzeit nicht die Anzeige von Text in Schreibrichtung von rechts nach links. Wenn Ihr Add-in jedoch  `fontDirection:` auf „rechts nach links" festlegt, wenn es in Excel Online ausgeführt wird, wird diese Formateinstellung in der Arbeitsmappendatei gespeichert und korrekt angezeigt, wenn die Arbeitsmappe in Desktop Excel geöffnet wird.|
| `fontStrikethrough:`|**Boolean**||
| `fontSuperscript:`|**Boolean**||
| `fontSubScript:`|**Boolean**||
| `fontNormal:`|**Boolean**|Legt die Schriftart, den Schriftschnitt, die Größe und Effekte zur normalen Formatierung fest. Dadurch wird das Zellenschriftformat auf die Standardwerte zurückgesetzt. Entspricht dem Aktivieren des Kontrollkästchens  **Standardschriftart** auf der Registerkarte **Schriftart** des Dialogfelds **Zellen formatieren**.|



**Einzug**


|**Schlüssel**|**Werte**|**Notizen**|
|:-----|:-----|:-----|
| `indentLeft:`|**Number**|Bei der Kombination mit einem Ausrichtungswert wird nur folgende Kombination unterstützt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">alignHorizontal: "left"</span> und <span class="code">indentLeft: <value></span></p></li></ul>|
| `indentRight:`|**Number**|Bei der Kombination mit einem Ausrichtungswert wird nur die folgende Kombination unterstützt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">alignHorizontal: "right"</span> und <span class="code">indentRight: <value></span></p></li></ul>|
| `indentDistributed:`|**Number**|Bei der Kombination mit einem Ausrichtungswert wird nur folgende Kombination unterstützt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">alignHorizontal: "distributed"</span> und <span class="code">indentDistributed: <value></span></p></li></ul>|



**Zahlenformatierung**


|**Schlüssel**|**Werte**|**Hinweise**|
|:-----|:-----|:-----|
| `numberFormat:`|**String**|Verwenden Sie zum Angeben der Zahlenformatierung eine benutzerdefinierte Zahlenformatzeichenfolge. Um beispielsweise zwei Dezimalstellen mit Komma als Tausendertrennzeichen anzugeben, müssten Sie so vorgehen:  `numberFormat:"#,###.00"`Hierbei handelt es sich um die gleichen benutzerdefinierten Formatzeichenfolgen, die Sie [mit der benutzerdefinierten Formatkategorie auf der Registerkarte "Zahlen" im Dialogfeld "Zellen formatieren" erstellen können](http://office.microsoft.com/de-de/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1). **Tipp:** Hier sehen Sie, wie eine Formatzeichenfolge für eine Standardkategorie im Dialogfeld **Zellen formatieren** in Excel aussieht:
<ol xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Wählen Sie eine Standardformatkategorie, z. B. <span class="ui">Währung</span>, aus der Liste <span class="ui">Kategorie</span> aus.</p></li><li><p>Legen Sie die Formatoptionen rechts im Dialogfeld fest.</p></li><li><p>Wählen Sie die Kategorie <span class="ui">Benutzerdefiniert</span> aus, um die Formatzeichenfolge oben in der Liste <span class="ui">Typ</span> anzuzeigen.</p></li></ol>|



**Tabellenoptionen**


|**Schlüssel**|**Werte**|**Hinweise**|
|:-----|:-----|:-----|
| `style:`|"keine" | <Alle vordefinierten Tabellenformatnamen>|Vordefinierte Tabellenformatnamen:"TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27", "TableStyleMedium28", "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11"Um zu sehen, wie ein Tabellenformat aussieht, fügen Sie eine Tabelle in Excel auf der Registerkarte  **Tabellentools** |**Entwurf** hinzu, klicken auf die Dropdownliste **Schnellformatvorlagen** und wählen anschließend ein vordefiniertes Format aus. Die QuickInfo des Formats entspricht einem der Werte aus der oben aufgeführten Liste.|
| `headerRow:`|**Boolean**||
| `firstColumn:`|**Boolean**||
| `filterButton:`|**Boolean**||
| `totalRow:`|**Boolean**||
| `lastColumn:`|**Boolean**||
| `bandedRows:`|**Boolean**||
| `bandedColumns:`|**Boolean**||
