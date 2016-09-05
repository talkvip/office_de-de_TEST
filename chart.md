# Diagrammobjekt (JavaScript-API für Excel)

Steht für ein Diagrammobjekt in einer Arbeitsmappe.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|height|double|Die Höhe des Diagrammobjekts (in Punkten).|
|id|string|Ruft ein Diagramm anhand seiner Position in der Sammlung ab. Schreibgeschützt.|
|left|double|Der Abstand von der linken Seite des Diagramms zu dem Ursprung des Arbeitsblatts (in Punkten).|
|name|string|Gibt den Namen eines Diagrammobjekts an.|
|top|double|Der Abstand (in Punkten) von dem oberen Rand des Objekts zum oberen Rand von Zeile 1 (auf einem Arbeitsblatt) oder zum oberen Diagrammbereich (in einem Diagramm).|
|width|double|Die Breite des Diagrammobjekts (in Punkten).|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Achsen|[ChartAxes](chartaxes.md)|Die Achsen des Diagramms. Schreibgeschützt.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Stellt die Datenbeschriftungen im Diagramm dar. Schreibgeschützt.|
|Format|[ChartAreaFormat](chartareaformat.md)|Kapselt die Formateigenschaften für den Diagrammbereich. Schreibgeschützt.|
|Legende|[ChartLegend](chartlegend.md)|Die Legende für das Diagramm. Schreibgeschützt.|
|Datenreihe|[ChartSeriesCollection](chartseriescollection.md)|Eine einzelne Datenreihe oder eine Sammlung von Datenreihen im Diagramm. Schreibgeschützt.|
|Titel|[ChartTitle](charttitle.md)|Der Titel des angegebenen Diagramms, einschließlich Text, Sichtbarkeit, Position und Formatierung des Titels. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Löscht das Diagrammobjekt.|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|System.IO.Stream|Rendert das Diagramm als base64-codiertes Bild durch Skalierung, um es an die angegebenen Maße anzupassen.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Setzt die Quelldaten für das Diagramm zurück.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positioniert das Diagramm im Verhältnis zu den Zellen im Arbeitsblatt.|

## Methodendetails


### delete()
Löscht das Diagrammobjekt.

#### Syntax
```js
chartObject.delete();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getImage(height: number, width: number, fittingMode: string)
Rendert das Diagramm als base64-codiertes Bild durch Skalierung, um es an die angegebenen Maße anzupassen.

#### Syntax
```js
chartObject.getImage(height, width, fittingMode);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|height|number|Optional. (Optional) Die gewünschte Höhe des resultierenden Bildes.|
|width|number|Optional. (Optional) Die gewünschte Breite des resultierenden Bildes.|
|fittingMode|string|Optional. (Optional) Die Methode, die verwendet wird, um das Diagramm auf die angegebenen Maße zu skalieren (wenn Höhe und Breite festgelegt sind).  Die folgenden Werte sind möglich: Fit, FitAndCenter, Fill|

#### Gibt Folgendes zurück:
System.IO.Stream

#### Beispiele
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var image = chart.getImage();
    return ctx.sync(); 
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

### setData(sourceData: Range, seriesBy: string)
Setzt die Quelldaten für das Diagramm zurück.

#### Syntax
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|sourceData|Range|Das den Quelldaten entsprechende Range-Objekt.|
|seriesBy|string|Optional. Gibt an, wie Spalten oder Zeilen als Datenreihen im Diagramm verwendet werden. Die folgenden Werte sind möglich: Auto, Columns, Rows. Auf dem Desktop prüft die Option „auto“ die Quelldatenform, um automatisch zu erkennen, ob die Daten in Zeilen oder Spalten vorliegen. In Excel Online ist der Wert für „auto“ standardmäßig „colums“.|

#### Gibt Folgendes zurück:
void

#### Beispiele

Legen Sie `sourceData` auf "A1: B4" fest und `seriesBy` auf "Spalten" fest.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### setPosition(startCell: Range or string, endCell: Range or string)
Positioniert das Diagramm im Verhältnis zu den Zellen im Arbeitsblatt.

#### Syntax
```js
chartObject.setPosition(startCell, endCell);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|startCell|Range or string|Die Startzelle. An diese Position wird das Diagramm verschoben. Die Startzelle ist die obere linke oder die obere rechte Zelle, abhängig davon, ob die eingestellte Textrichtung des Benutzers von links nach rechts oder von rechts nach links ist.|
|endCell|Range or string|Optional. Die Endzelle. Wenn angegeben, werden Breite und Höhe des Diagramms so eingestellt, dass diese Zelle/dieser Bereich vollständig bedeckt ist.|

#### Gibt 
void

#### Beispiele


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### Eigenschaftszugriffsbeispiele

Abrufen eines Diagramms mit dem Namen "Diagramm1"

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.load('name');
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

Aktualisieren Sie ein Diagramm, indem Sie es umbenennen, positionieren und die Größe anpassen.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.weight = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Geben Sie dem Diagramm einen neuen Namen, legen Sie die Höhe und Stärke des Diagramms auf 200 Punkte fest. Verschieben Sie Diagramm1 auf 100 Punkte nach oben und nach links. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";  
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

