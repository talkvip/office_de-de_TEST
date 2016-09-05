# ImageOcrData-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt Daten dar, die von der OCR (optische Zeichenerkennung) eines Bilds abgerufen werden.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|ocrLanguageId|string|Stell die OCR-Sprache mit Werten, wie z. B. EN-US, dar.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrLanguageId)|
|ocrText|string|Stelllt den Text dar, der von der OCR des Bilds abgerufen wird.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrText)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-load)|

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
**ocrText und ocrLanguageId**
```js
var image = null;

OneNote.run(function(ctx){
    // Get the current outline.
    var outline = ctx.application.getActiveOutline();

    // Queue a command to load paragraphs and their types.
    outline.load("paragraphs")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
            if (image != null)
            {
               image.load("ocrData");
            }
            return ctx.sync();
        })
        .then(function(){
            
            // Log ocrText and ocrLanguageId
            console.log(image.ocrData.ocrText);
            console.log(image.ocrData.ocrLanguageId);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
