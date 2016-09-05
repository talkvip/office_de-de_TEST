# RichText-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt ein RichText-Objekt in einem Absatz dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Ruft die ID des RichText-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-id)|
|text|string|Ruft den Textinhalt des RichText-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-text)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraph|[Absätzen](paragraph.md)|Ruft das Paragraph-Objekt ab, das das RichText-Objekt enthält. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-paragraph)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-load)|

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

**id und text**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its outline's paragraphs.
    var outlinePageContents = [];
    var paragraphs = [];
    var richTextParagraphs = [];
    // Queue a command to load the id and type of each page content in the outline.
    pageContents.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Load all page contents of type Outline
            $.each(pageContents.items, function(index, pageContent) {
                if(pageContent.type == 'Outline')
                {
                    pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                    outlinePageContents.push(pageContent);
                }
            });
            return context.sync();
        })
        .then(function () {
            // Load all rich text paragraphs across outlines
            $.each(outlinePageContents, function(index, outlinePageContent) {
                var outline = outlinePageContent.outline;
                paragraphs = paragraphs.concat(outline.paragraphs.items);
            });
            $.each(paragraphs, function(index, paragraph) {
                if(paragraph.type == 'RichText')
                {
                    richTextParagraphs.push(paragraph);
                    paragraph.load("id,richText/text");
                }
            });
            return context.sync();
        })
        .then(function () {
            // Display all rich text paragraphs to the console
            $.each(richTextParagraphs, function(index, richTextParagraph) {
                var richText = richTextParagraph.richText;
                console.log("Paragraph found with richtext content : " + richText.text + " and richtext id : " + richText.id);
            });
            return context.sync();
        });
});
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
