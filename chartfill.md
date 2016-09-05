# ChartFill-Objekt (JavaScript-API für Excel)

Die Füllung für ein Diagrammelement.

## Eigenschaften

Keine

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Entfernen der Füllfarbe eines Diagrammelements|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Legt die Füllung eines Diagrammelements auf einfarbige Füllung fest.|

## Methodendetails


### clear()
Entfernen der Füllfarbe eines Diagrammelements

#### Syntax
```js
chartFillObject.clear();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

Löschen der Linienformatierung der Hauptgitternetzlinien auf der Größenachse des Diagramms mit dem Namen „Diagramm1“

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### setSolidColor(color: string)
Legt die Füllung eines Diagrammelements auf einfarbige Füllung fest.

#### Syntax
```js
chartFillObject.setSolidColor(color);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|color|string|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B. „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|

#### Returns
void

#### Beispiele

Festlegen der Hintergrundfarbe in Diagramm1 auf Rot.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
