# Worksheet-Objekt (JavaScript-API für Excel)

Ein Excel-Arbeitsblatt ist ein Raster von Zellen. Es kann Daten, Tabellen, Diagramme usw. enthalten.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|id|string|Gibt einen Wert zurück, der das Arbeitsblatt in einer bestimmten Arbeitsmappe eindeutig identifiziert. Der Wert des Bezeichners bleibt unverändert, auch wenn das Arbeitsblatt umbenannt oder verschoben wird. Diese Werte ändern sich mit jeder Sitzung, in der die Datei geöffnet wird. Schreibgeschützt.|
|name|string|Der Anzeigename des Arbeitsblatts.|
|position|int|Die nullbasiert Position des Arbeitsblatts in der Arbeitsmappe.|
|visibility|string|Die Sichtbarkeit des Arbeitsblatts. Die folgenden Werte sind möglich: Visible, Hidden, VeryHidden.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Diagramme|[ChartCollection](chartcollection.md)|Gibt die Sammlung von Diagrammen zurück, die Teil des Arbeitsblatts sind. Schreibgeschützt.|
|Schutz|[WorksheetProtection](worksheetprotection.md)|Gibt das Arbeitsblattschutz-Objekt für ein Arbeitsblatt zurück. Schreibgeschützt.|
|Tabellen|[TableCollection](tablecollection.md)|Gibt die Sammlung von Tabellen zurück, die Teil des Arbeitsblatts sind. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Aktivieren Sie das Arbeitsblatt in der Excel-Benutzeroberfläche.|
|[delete()](#delete)|void|Löscht das Arbeitsblatt aus der Arbeitsmappe.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält. Die Zelle kann sich außerhalb seines übergeordneten Bereichs befinden, solange es im Arbeitsblatt-Raster bleibt.|
|[getRange(address: string)](#getrangeaddress-string)|[Bereichen](range.md)|Ruft das durch die Adresse oder den Namen angegebene Bereichsobjekt ab.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Bereichen](range.md)|Der verwendete Bereich ist der kleinste Bereich, der mindestens eine der Zellen umfasst, die einen Wert enthalten oder denen eine Formatierung zugewiesen wurde. Wenn das Arbeitsblatt leer ist, gibt diese Funktion die oberste linke Zelle zurück.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### activate()
Aktivieren Sie das Arbeitsblatt in der Excel-Benutzeroberfläche.

#### Syntax
```js
worksheetObject.activate();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete()
Löscht das Arbeitsblatt aus der Arbeitsmappe.

#### Syntax
```js
worksheetObject.delete();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält. Die Zelle kann sich außerhalb seines übergeordneten Bereichs befinden, solange es im Arbeitsblatt-Raster bleibt.

#### Syntax
```js
worksheetObject.getCell(row, column);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|row|number|Die Zeilenanzahl der abzurufenden Zelle. Nullindiziert.|
|column|number|Die Spaltenanzahl der abzurufenden Zelle. Nullindiziert.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
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


### getRange(address: string)
Ruft das durch die Adresse oder den Namen angegebene Bereichsobjekt ab.

#### Syntax
```js
worksheetObject.getRange(address);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|address|string|Optional. Die Adresse oder der Name des Bereichs. Wenn nichts angegeben ist, wird der gesamte Arbeitsblattbereich zurückgegeben.|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
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
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getUsedRange(valuesOnly: bool)
Der verwendete Bereich ist der kleinste Bereich, der mindestens eine der Zellen umfasst, die einen Wert enthalten oder denen eine Formatierung zugewiesen wurde. Wenn das Arbeitsblatt leer ist, gibt diese Funktion die oberste linke Zelle zurück.

#### Syntax
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Betrachtet nur Zellen mit Werten als verwendet (ignoriert die Formatierung).|

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
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
### Eigenschaftszugriffsbeispiele

Rufen Sie die Arbeitsblatteigenschaften auf Grundlage des Blattnamens ab.

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Legen Sie die Arbeitsblattposition fest. 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

