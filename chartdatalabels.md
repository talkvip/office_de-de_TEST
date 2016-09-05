# ChartDataLabels-Objekt (JavaScript-API für Excel)

Stellt eine Sammlung aller Datenbeschriftungen an einem Diagrammpunkt dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Position|string|DataLabelPosition-Wert, der die Position der Datenbeschriftung darstellt. Die folgenden Werte sind möglich: Keine, Mitte, InsideEnd, InsideBase, OutsideEnd, links, rechts, oben, unten, BestFit, Callout. Schreibzugriff.|
|Separator|string|Zeichenfolge, die das Trennzeichen für die Datenbeschriftungen in einem Diagramm darstellt. Schreibzugriff.|
|showBubbleSize|bool|Boolescher Wert, der angibt, ob die Größe der Datenbeschriftungs-Sprechblase angezeigt wird. Schreibzugriff.|
|showCategoryName|bool|Boolescher Wert, der angibt, ob der Kategoriename der Datenbeschriftung angezeigt wird. Schreibzugriff.|
|showLegendKey|bool|Boolescher Wert, der angibt, ob das Legendensymbol der Datenbeschriftung angezeigt wird. Schreibzugriff.|
|showPercentage|bool|Boolescher Wert, der angibt, ob der Prozentsatz der Datenbeschriftung angezeigt wird. Schreibzugriff.|
|showSeriesName|bool|Boolescher Wert, der angibt, ob der Name der Datenbeschriftungsreihe angezeigt wird. Schreibzugriff.|
|showValue|bool|Boolescher Wert, der angibt, ob der Datenbeschriftungswert angezeigt wird. Schreibzugriff.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Format|[ChartDataLabelFormat](chartdatalabelformat.md)|Stellt das Format der Diagrammdatenbeschriftungen dar, einschließlich Füllung und Formatierung der Schriftart. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


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

Stellen Sie sicher, dass Namen der Datenbeschriftungsreihe angezeigt werden, und legen Sie `position` der Datenbeschriftungen auf „top“ fest.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.visible = true;
    chart.datalabels.position = "top";
    chart.datalabels.ShowSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
