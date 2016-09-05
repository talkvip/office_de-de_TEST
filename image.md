# Image-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt ein Bild dar. Ein Bild kann ein direkt untergeordnetes Element eines PageContent-Objekts oder eines Paragraph-Objekts sein.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|
            description|string|Dient zum Abrufen oder Festlegen der Beschreibung des Bilds.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-description)|
|Höhe|double|Dient zum Abrufen oder Festlegen der Höhe des Bildlayouts.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-height)|
|Hyperlink|string|Dient zum Abrufen oder Festlegen des Hyperlinks des Bilds.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-hyperlink)|
|id|string|Ruft die ID des Image-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-id)|
|Breite|double|Dient zum Abrufen oder Festlegen der Breite des Bildlayouts.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-width)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|ocrData|[ImageOcrData](imageocrdata.md)|Ruft die Daten ab, die von der optischen Zeichenerkennung (OCR) für dieses Bild abgerufen werden, z. B. OCR-Text und Sprache. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-ocrData)|
|pageContent|[PageContent](pagecontent.md)|Ruft das PageContent-Objekt ab, das das Bild enthält. Wird ausgelöst, wenn das Bild nicht ein direkt untergeordnetes Element eines PageContent-Objekts ist. Dieses Objekt definiert die Position des Bilds auf der Seite. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-pageContent)|
|Absatz|[Absätzen](paragraph.md)|Ruft das Paragraph-Objekt ab, das das Bild enthält. Wird ausgelöst, wenn das Bild nicht ein direkt untergeordnetes Element eines Absatzes ist. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-paragraph)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getBase64Image()](#getbase64image)|string|Ruft die base64-verschlüsselte binäre Darstellung des Bilds ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-getBase64Image)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-load)|

## Methodendetails


### getBase64Image()
Ruft die base64-verschlüsselte binäre Darstellung des Bilds ab.

#### Syntax
```js
imageObject.getBase64Image();
```

#### Parameter
Keine

#### Gibt 
string

#### Beispiele
```js

var image = null;
var imageString;

OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
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
        })
        .then(function(){
            if (image != null)
            {
                imageString = image.getBase64Image();
                return ctx.sync();
            }
        })
        .then(function(){
            console.log(imageString);
        });
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
**ID, Breite, Höhe, Beschreibung und Link**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var image = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
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
        })
        .then(function(){
            if (image != null)
            {
                // load every properties and relationships
                ctx.load(image);
                return ctx.sync();
            }
        })
        .then(function(){
            if (image != null)
            {                   
                console.log("image " + image.id + " width is " + image.width + " height is " + image.height);
                console.log("description: " + image.description);                   
                console.log("hyperlink: " + image.hyperlink);
            }
        });
});
```

**ocrData**
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
            console.log(image.ocrData);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**Absatz**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var searchedParagraph = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function() {
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    searchedParagraph = paragraph;
                    break;
                }
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {
                // load every properties and relationships
                searchedParagraph.image.load('paragraph');
                return ctx.sync();
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {                   
                if (searchedParagraph.id != searchedParagraph.image.paragraph.id)
                {
                    console.log("id must match");
                }
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

