# ChartPointsCollection-Objekt (JavaScript-API für Excel)

Eine Sammlung aller Diagrammpunkte in einer Datenreihe eines Diagramms.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|count|int|Gibt die Anzahl der Diagrammpunkte in der Sammlung zurück. Schreibgeschützt.|
|Items|[ChartPoint[]](chartpoint.md)|Eine Sammlung von chartPoints-Objekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|Abrufen eines Punkts anhand seiner Position in der Datenreihe.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getItemAt(index: number)
Abrufen eines Punkts anhand seiner Position in der Datenreihe.

#### Syntax
```js
chartPointsCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[ChartPoint](chartpoint.md)

#### Beispiele
Festlegen der Rahmenfarbe für die ersten Punkte in der Punktesammlung

```js
Excel.run(function (ctx) { 
    var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    points.getItemAt(0).format.fill.setSolidColor("#8FBC8F");
    return ctx.sync().then(function() {
        console.log("Point Border Color Changed");
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

Abrufen der Namen von Punkten in der Punktesammlung

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
    pointsCollection.load('items');
    return ctx.sync().then(function() {
        console.log("Points Collection loaded");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Abrufen der Anzahl von Punkten

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
    pointsCollection.load('count');
    return ctx.sync().then(function() {
        console.log("points: Count= " + pointsCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
