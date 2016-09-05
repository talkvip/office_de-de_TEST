# Bereichsobjekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Ein Bereich stellt einen Satz von einer oder mehrerer zusammenhängender Zellen wie z. B. eine Zelle, eine Zeile oder eine Spalte, ein Block von Zellen usw. dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|address|string|Entspricht dem Bereichsverweis in A1-Schreibweise. Adresswert für den Blattbezug (z. B. Tabelle1!A1:B4). Schreibgeschützt.|
|addressLocal|string|Stellt den Bereichsbezug für den angegebenen Bereich in der Sprache des Benutzers dar. Schreibgeschützt.|
|cellCount|int|Anzahl der Zellen im Bereich. Schreibgeschützt.|
|columnCount|int|Stellt die Gesamtanzahl der Spalten im Bereich dar. Schreibgeschützt.|
|columnHidden|bool|Stellt dar, ob alle Spalten des aktuellen Bereichs ausgeblendet sind.|
|columnIndex|int|Stellt die Spaltenanzahl der ersten Zelle im Bereich dar. Nullindiziert. Schreibgeschützt.|
|formulas|object[]|Stellt die Formel in der A1-Schreibweise dar.|
|formulasLocal|object[][]|Stellt die Formel in der A1-Schreibweise, Sprache des Benutzers und im Gebietsschema der Zahlenformatierung dar.  Beispielsweise würde die englische Formel „= SUM(A1, 1.5)“ in Deutsch „= SUMME(A1; 1,5)“ werden.|
|formulasR1C1|object[][]|Stellt die Formel in der R1C1-Schreibweise dar.|
|hidden|bool|Stellt dar, ob alle Zellen des aktuellen Bereichs ausgeblendet sind. Schreibgeschützt.|
|numberFormat|object[][]|Stellt den Zahlenformatcode für die angegebene Zelle dar.|
|rowCount|int|Gibt die Anzahl der Zeilen im Bereich zurück. Schreibgeschützt.|
|rowHidden|bool|Stellt dar, ob alle Zeilen des aktuellen Bereichs ausgeblendet sind.|
|rowIndex|int|Gibt die Spaltenanzahl der ersten Zelle im Bereich zurück. Nullindiziert. Schreibgeschützt.|
|text|object[][]|Textwerte des angegebenen Bereichs. Der Textwert hängt nicht von der Zellenbreite ab. Die Ersetzung des #-Zeichens, die in der Excel-Benutzeroberfläche passiert, wirkt sich nicht auf den von der API zurückgegebenen Textwert aus. Schreibgeschützt.|
|valueTypes|string|Stellt den Datentyp in jeder Zelle dar. Schreibgeschützt. Die folgenden Werte sind möglich: Unbekannt, leer, Zeichenfolge, Ganzzahl, Doppelwort, Boolesch, Fehler.|
|values|object[][]|Stellt die Rohwerte des angegebenen Bereichs dar. Die zurückgegebenen Daten können den Typ Zeichenfolge, Zahl oder ein boolescher Wert sein. Eine Zelle, die einen Fehler enthält, gibt eine Fehlerzeichenfolge zurück.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Format|[RangeFormat](rangeformat.md)|Gibt ein Formatobjekt zurück, das die Schriftart des Bereichs, Füllung, den Rahmen, die Ausrichtung und andere Eigenschaften verschachtelt. Schreibgeschützt.|
|sort|[RangeSort](rangesort.md)|Stellt die Sortierkonfiguration für den Bereich dar. Schreibgeschützt.|
|worksheet|[Arbeitsblatt](worksheet.md)|Das Arbeitsblatt, das den aktuellen Bereich enthält. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|Löschen von Bereichswerten, Format, Füllung, Rahmen usw.|
|[delete(shift: string)](#deleteshift-string)|void|Löscht die einem Bereich zugeordneten Zellen.|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Bereichen](range.md)|Ruft das kleinste Bereichsobjekt ab, das die angegebenen Bereiche umfasst. Beispielsweise das getBoundingRect von "B2:C5" und "D10:E15" lautet "B2:E15".|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält. Die Zelle kann sich außerhalb seines übergeordneten Bereichs befinden, solange es im Arbeitsblatt-Raster bleibt. Die zurückgegebene Zelle befindet sich relativ zur obersten linken Zelle des Bereichs.|
|[getColumn(column: number)](#getcolumncolumn-number)|[Bereichen](range.md)|Ruft eine Spalte ab, die im Bereich enthalten ist.|
|[getEntireColumn()](#getentirecolumn)|[Bereichen](range.md)|Ruft ein Objekt ab, das die gesamte Spalte des Bereichs darstellt.|
|[getEntireRow()](#getentirerow)|[Bereichen](range.md)|Ruft ein Objekt ab, das die gesamte Zeile des Bereichs darstellt.|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt.|
|[getLastCell()](#getlastcell)|[Bereichen](range.md)|Ruft die letzte Zelle im Bereich ab. Beispielsweise lautet die letzte Zelle des Bereichs „B2: D5“ „D5“.|
|[getLastColumn()](#getlastcolumn)|[Bereichen](range.md)|Ruft die letzte Spalte im Bereich ab. Beispielsweise lautet die letzte Spalte von „B2:D5“ „D2:D5“.|
|[getLastRow()](#getlastrow)|[Bereichen](range.md)|Ruft die letzte Zeile im Bereich ab. Beispielsweise lautet die letzte Zelle des Bereichs "B2: D5" "B5:D5".|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Bereichen](range.md)|Ruft ein Objekt ab, das einen Bereich darstellt, der aus dem angegebenen Bereich versetzt ist. Die Dimension des zurückgegebenen Bereichs entspricht diesem Bereich. Wenn der resultierende Bereich außerhalb des Arbeitsblatt-Rasters erzwungen wird, wird eine Ausnahme ausgelöst.|
|[getRow(row: number)](#getrowrow-number)|[Bereichen](range.md)|Ruft eine Zelle ab, die im Bereich enthalten ist.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Bereichen](range.md)|Gibt den verwendeten Unterbereich des Bereichsobjekts zurück.|
|[insert(shift: string)](#insertshift-string)|[Bereichen](range.md)|Fügt eine Zelle oder einen Zellbereich in das Arbeitsblatt anstelle dieses Bereichs ein, und verschiebt die anderen Zellen, um Platz zu schaffen. Gibt ein neues Bereichsobjekt in dem nun leeren Bereich zurück.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|
|[merge(across: bool)](#mergeacross-bool)|void|Führt die Zellen des Bereichs in einen Bereich im Arbeitsblatt zusammen.|
|[select](#select)|void|Wählt den angegebenen Bereich in der Excel-Benutzeroberfläche aus.|
|[unmerge()](#unmerge)|void|Hebt den Zellverbund des Bereichs in einzelne Zellen auf.|

## Methodendetails


### clear(applyTo: string)
Löschen von Bereichswerten, Format, Füllung, Rahmen usw.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Bestimmt den Typ der Löschaktion. Die folgenden Werte sind möglich: `All` Standardoption,`Formats` ,`Contents`|

#### Gibt 
void

#### Beispiele

Im Beispiel unten wird das Format und der Inhalt des Bereichs gelöscht. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete(shift: string)
Löscht die einem Bereich zugeordneten Zellen.

#### Syntax
```js
rangeObject.delete(shift);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Shift|string|Gibt an, wohin die Zellen verschoben werden.  Die folgenden Werte sind möglich: Nach oben, nach links|

#### Gibt 
void

#### Beispiele

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getBoundingRect(anotherRange: Range or string)
Ruft das kleinste Bereichsobjekt ab, das die angegebenen Bereiche umfasst. Beispielsweise das GetBoundingRect von "B2:C5" und "D10:E15" lautet "B2:E15".

#### Syntax
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|anotherRange|Bereich oder Zeichenfolge|Das Bereichsobjekt oder die Adresse oder der Bereichsname.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält. Die Zelle kann sich außerhalb seines übergeordneten Bereichs befinden, solange es im Arbeitsblatt-Raster bleibt. Die zurückgegebene Zelle befindet sich relativ zur obersten linken Zelle des Bereichs.

#### Syntax
```js
rangeObject.getCell(row, column);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|row|number|Zeilenanzahl der abzurufenden Zelle. Nullindiziert.|
|column|number|Spaltenanzahl der abzurufenden Zelle. Nullindiziert.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getColumn(column: number)
Ruft eine Spalte ab, die im Bereich enthalten ist.

#### Syntax
```js
rangeObject.getColumn(column);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|column|number|Spaltenanzahl des abzurufenden Bereichs. Nullindiziert.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getEntireColumn()
Ruft ein Objekt ab, das die gesamte Spalte des Bereichs darstellt.

#### Syntax
```js
rangeObject.getEntireColumn();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

Hinweis: Die Rastereigenschaften des Bereichs (Werte, NumberFormat, Formeln) enthalten `null`, da der betroffene Bereich ungebunden ist.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getEntireRow()
Ruft ein Objekt ab, das die gesamte Zeile des Bereichs darstellt.

#### Syntax
```js
rangeObject.getEntireRow();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Die Rastereigenschaften des Bereichs (Werte, NumberFormat, Formeln) enthalten `null`, da der betroffene Bereich ungebunden ist.

### getIntersection(anotherRange: Range or string)
Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt.

#### Syntax
```js
rangeObject.getIntersection(anotherRange);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|anotherRange|Bereich oder Zeichenfolge|Das Bereichsobjekt oder die Bereichsadresse, die verwendet wird, um die Schnittmenge der Bereiche zu ermitteln.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastCell()
Ruft die letzte Zelle im Bereich ab. Beispielsweise lautet die letzte Zelle des Bereichs „B2: D5“ „D5“.

#### Syntax
```js
rangeObject.getLastCell();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastColumn()
Ruft die letzte Spalte im Bereich ab. Beispielsweise lautet die letzte Spalte von „B2:D5“ „D2:D5“.

#### Syntax
```js
rangeObject.getLastColumn();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastRow()
Ruft die letzte Zeile im Bereich ab. Beispielsweise lautet die letzte Zelle des Bereichs "B2: D5" "B5:D5".

#### Syntax
```js
rangeObject.getLastRow();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### getOffsetRange(rowOffset: number, columnOffset: number)
Ruft ein Objekt ab, das einen Bereich darstellt, der aus dem angegebenen Bereich versetzt ist. Die Dimension des zurückgegebenen Bereichs entspricht diesem Bereich. Wenn der resultierende Bereich außerhalb des Arbeitsblatt-Rasters erzwungen wird, wird eine Ausnahme ausgelöst.

#### Syntax
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|rowOffset|number|Die Anzahl an Zeilen (positiv, negativ oder 0), um die der Bereich versetzt werden muss. Bei positiven Werten wird der Bereich nach unten versetzt. Bei negativen Werten wird der Bereich nach oben versetzt.|
|columnOffset|number|Die Anzahl an Spalten (positiv, negativ oder 0), um die der Bereich versetzt werden muss. Bei positiven Werten wird der Bereich nach unten versetzt. Bei negativen Werten wird der Bereich nach oben versetzt.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRow(row: number)
Ruft eine Zelle ab, die im Bereich enthalten ist.

#### Syntax
```js
rangeObject.getRow(row);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|row|number|Zeilenanzahl des abzurufenden Bereichs. Nullindiziert.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getUsedRange(valuesOnly: bool)
Gibt den verwendeten Bereich des angegebenen Bereichsobjekts zurück.

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Wenn „true“, dann werden nur Zellen, die momentan Werte aufweisen, als verwendet betrachtet. Die Standardeinstellung „false“ zählt jede Zelle, die schon einmal einen Wert aufgewiesen hat, als verwendet.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### insert(shift: string)
Fügt eine Zelle oder einen Zellbereich in das Arbeitsblatt anstelle dieses Bereichs ein, und verschiebt die anderen Zellen, um Platz zu schaffen. Gibt ein neues Bereichsobjekt in dem nun leeren Bereich zurück.

#### Syntax
```js
rangeObject.insert(shift);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Shift|string|Gibt an, wohin die Zellen verschoben werden.  Die folgenden Werte sind möglich: Nach unten, nach rechts|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### load(param: object)
Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.

#### Syntax
```js
object.load(param);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder geben Sie das [loadOption](loadoption.md)-Objekt an.|

#### Gibt 
void

### merge(across: bool)
Führt die Zellen des Bereichs in einen Bereich im Arbeitsblatt zusammen.

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|across|bool|Optional. Mit „true“ werden Zellen in allen Zeilen des angegebenen Bereichs als einzelne zusammengeführte Zellen zusammengeführt. Der Standardwert lautet „false“.|

#### Returns
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### select()
Wählt den angegebenen Bereich in der Excel-Benutzeroberfläche aus.

#### Syntax
```js
rangeObject.select();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### unmerge()
Hebt den Zellverbund des Bereichs in einzelne Zellen auf.

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### Eigenschaftszugriffsbeispiele

In diesem Beispiel wird die Bereichsadresse zum Abrufen des Bereichsobjekts verwendet.

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8"; 
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

In diesem Beispiel wird ein benannter Bereich zum Abrufen des Bereichsobjekts verwendet.

```js

Excel.run(function (ctx) { 
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Im untenstehenden Beispiel werden das Zahlenformat, Werte und Formeln auf einem Raster festgelegt, das ein 2x3-Raster enthält.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Das Beispiel unten entspricht dem oben, mit der Ausnahme, dass es die R1C1-Schreibweise für die Formeln verwendet.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulasR1C1 = [[null,null], [null,null], [null,"=R[-1]C-R[-2]C"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulasR1C1= formulasR1C1;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Rufen Sie das Arbeitsblatt ab, das den Bereich enthält. 

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

