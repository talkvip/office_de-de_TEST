# ChartCollection-Objekt (JavaScript-API für Excel)

Eine Sammlung aller Diagrammobjekte in einem Arbeitsblatt.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|count|int|Gibt die Anzahl der Diagramme im Arbeitsblatt zurück. Schreibgeschützt.|
|Items|[Chart[]](chart.md)|Eine Sammlung von Diagrammobjekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Diagramm](chart.md)|Erstellt ein neues Diagramm.|
|[getItem(name: string)](#getitemname-string)|[Diagramm](chart.md)|Ruft ein Diagramm über seinen Namen ab. Wenn mehrere Diagramme mit demselben Namen vorhanden sind, wird das erste zurückgegeben.|
|[getItemAt(index: number)](#getitematindex-number)|[Diagramm](chart.md)|Ruft ein Diagramm anhand seiner Position in der Sammlung ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### add(type: string, sourceData: Range, seriesBy: string)
Erstellt ein neues Diagramm.

#### Syntax
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|type|string|Stellt die Art des Diagramms dar.  Die folgenden Werte sind möglich: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie usw.|
|sourceData|Range|Das Bereichsobjekt, in dem die Quelldaten enthalten sind.|
|seriesBy|string|Optional. Gibt an, wie Spalten oder Zeilen als Datenreihen im Diagramm verwendet werden.  Die folgenden Werte sind möglich: Automatisch, Spalten, Zeilen|

#### Gibt Folgendes zurück:
[Diagramm](chart.md)

#### Beispiele

Hinzufügen eines Diagramms vom `chartType` „ColumnClustered“ im Arbeitsblatt „Diagramme“ mit `sourceData` aus dem Bereich „A1: B4“ und Festlegen von `seriesBy` auf „automatisch“.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(name: string)
Ruft ein Diagramm über seinen Namen ab. Wenn mehrere Diagramme mit demselben Namen vorhanden sind, wird das erste zurückgegeben.

#### Syntax
```js
chartCollectionObject.getItem(name);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|name|string|Der Name des abzurufenden Diagramms.|

#### Gibt Folgendes zurück:
[Diagramm](chart.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### Beispiele

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### Beispiele

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Ruft ein Diagramm anhand seiner Position in der Sammlung ab.

#### Syntax
```js
chartCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[Diagramm](chart.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
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

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
            console.log(charts.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Abrufen der Anzahl von Diagrammen

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

