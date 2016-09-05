# RangeFormat-Objekt (JavaScript-API für Excel)

Gibt ein Formatobjekt zurück, das die Schriftart des Bereichs, Füllung, Rahmen, die Ausrichtung und andere Eigenschaften verschachtelt.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|columnWidth|double|Ruft die Breite aller Spalten innerhalb des Bereichs ab oder legt diese fest. Wenn die Breite der Spalten nicht gleichmäßig ist, wird Null zurückgegeben.|
|horizontalAlignment|string|Stellt die horizontale Ausrichtung für das angegebene Objekt dar. Die folgenden Werte sind möglich: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|
|rowHeight|double|Ruft die Höhe aller Zeilen des Bereichs ab oder legt diese fest. Wenn die Höhe der Zeilen nicht gleichmäßig ist, wird Null zurückgegeben.|
|verticalAlignment|string|Stellt die vertikale Ausrichtung für das angegebene Objekt dar. Die folgenden Werte sind möglich: Top, Center, Bottom, Justify, Distributed.|
|wrapText|bool|Gibt an, dass das Excel-Textsteuerelement auf das Umbrechen von Text im Objekt festgelegt ist. Ein Nullwert gibt an, dass der gesamte Bereich keine einheitliche Textumbruch-Einstellung verwendet.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Borders|[RangeBorderCollection](rangebordercollection.md)|Auflistung von Border-Objekten, die für den gesamten ausgewählten Bereich gelten, schreibgeschützt.|
|fill|[RangeFill](rangefill.md)|Gibt das Fill-Objekt an, das für den gesamten Bereich definiert ist. Schreibgeschützt.|
|font|[RangeFont](rangefont.md)|Gibt das Font-Objekt zurück, das für den gesamten ausgewählten Bereich definiert ist, schreibgeschützt.|
|Schutz|[FormatProtection](formatprotection.md)|Gibt das Formatschutz-Objekt für einen Bereich zurück. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|Ändert die Breite der Spalten des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Spalten die optimale Breite zu erzielen.|
|[autofitRows()](#autofitrows)|void|Ändert die Höhe der Zeilen des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Zeilen die optimale Höhe zu erzielen.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### autofitColumns()
Ändert die Breite der Spalten des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Spalten die optimale Breite zu erzielen.

#### Syntax
```js
rangeFormatObject.autofitColumns();
```

#### Parameter
Keine

#### Gibt 
void

### autofitRows()
Ändert die Höhe der Zeilen des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Zeilen die optimale Höhe zu erzielen.

#### Syntax
```js
rangeFormatObject.autofitRows();
```

#### Parameter
Keine

#### Gibt 
void

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

Dieses Beispiel gibt alle Formateigenschaften eines Bereichs aus. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Im folgenden Beispiel wird der Schriftartenname und die Füllfarbe eines Bereichs festgelegt und der Text umgebrochen. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Im folgenden Beispiel wird um den Bereich ein Raster-Rahmen hinzugefügt.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
    range.format.borders('InsideVertical').lineStyle = 'Continuous';
    range.format.borders('EdgeBottom').lineStyle = 'Continuous';
    range.format.borders('EdgeLeft').lineStyle = 'Continuous';
    range.format.borders('EdgeRight').lineStyle = 'Continuous';
    range.format.borders('EdgeTop').lineStyle = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
