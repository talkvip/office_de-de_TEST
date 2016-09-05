# ChartLineFormat-Objekt (JavaScript-API für Excel)

Kapselt die Formatierungsoptionen für Linienelemente.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|color|string|HTML-Farbcode, der die Farbe der Linien im Diagramm darstellt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Löschen der Linienformatierung eines Diagrammelements.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### clear()
Löschen der Linienformatierung eines Diagrammelements.

#### Syntax
```js
chartLineFormatObject.clear();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

Löschen der Linienformatierung der Hauptgitternetzlinien für die Größenachse des Diagramms mit dem Namen „Diagramm1“

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

Festlegen der Hauptgitternetzlinien des Diagramms für die Größenachse auf Rot

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function() {
            console.log("Chart Gridlines Color Updated");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
