# ChartFont-Objekt (JavaScript-API für Excel)

Dieses Objekt stellt die Zeichenformatierung (Schriftart, Schriftgrad, Farbe usw.) für ein Diagrammobjekt dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|bold|bool|Stellt die Einstellung „Fett“ für die Schriftart dar.|
|color|string|HTML-Farbcodedarstellung der Textfarbe. #FF0000 stellt beispielsweise Rot dar.|
|italic|bool|Stellt die Einstellung „Kursiv“ für die Schriftart dar.|
|name|string|Name der Schriftart, z. B. Calibri.|
|size|double|Die Größe der Schriftart, z. B. 11.|
|underline|string|Art der auf die Schriftart angewendete Unterstreichung. Die folgenden Werte sind möglich: Keine, einfache.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


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

Verwenden Sie den Diagrammtitel als Beispiel.

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Legen Sie das Diagrammtitelformat auf Calibri, Schriftgrad 10, Fett und Rot fest. 

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
