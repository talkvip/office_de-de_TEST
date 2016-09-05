# ChartSeriesCollection-Objekt (JavaScript-API für Excel)

Stellt eine Sammlung von Diagrammdatenreihen dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|count|int|Gibt die Anzahl der Datenreihen in der Sammlung zurück. Schreibgeschützt.|
|Items|[ChartSeries[]](chartseries.md)|Eine Sammlung von chartSeries-Objekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Ruft eine Datenreihe anhand ihrer Position in der Sammlung ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getItemAt(index: number)
Ruft eine Datenreihe anhand ihrer Position in der Sammlung ab.

#### Syntax
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[ChartSeries](chartseries.md)

#### Beispiele

Abrufen des Namens der ersten Datenreihe in der Datenreihensammlung

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        console.log(seriesCollection.items[0].name);
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
Abrufen der Namen der Datenreihen in der Datenreihensammlung

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < seriesCollection.items.length; i++)
        {
            console.log(seriesCollection.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Abrufen der Anzahl von Diagrammdatenreihen in der Sammlung

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('count');
    return ctx.sync().then(function() {
        console.log("series: Count= " + seriesCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

