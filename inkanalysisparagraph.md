# InkAnalysisParagraph-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt Freihandanalysedaten für einen identifizierten Absatz bestehend aus Freihandstrichen dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Ruft die ID des InkAnalysisParagraph-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-id)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|inkAnalysis|[InkAnalysis](inkanalysis.md)|Verweis auf das übergeordnete InkAnalysisPage-Element. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-inkAnalysis)|
|Zeilen|[InkAnalysisLineCollection](inkanalysislinecollection.md)|Ruft die Freihandanalysezeilen in diesem  Freihandanalyseabsatz ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-lines)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-load)|

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

**Zeilen**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load a line of ink words.
    page.load('inkAnalysisOrNull/paragraphs/lines');
    
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            
            // Log id of each line in ink paragraphs.
            $.each(inkParagraphs.items, function(i, inkParagraph){
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function (j, inkLine) {
                    console.log(inkLine.id);
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```