# PageContentCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt den Inhalt einer Seite als eine Auflistung von PageContent-Objekten dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl der Seiteninhalten in der Auflistung zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-count)|
|items|[PageContent[]](pagecontent.md)|Eine Auflistung von pageContent-Objekten. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[PageContent](pagecontent.md)|Ruft ein PageContent-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|Ruft einen Seiteninhalt anhand seiner Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-load)|

## Methodendetails


### getItem(index: number or string)
Ruft ein PageContent-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
pageContentCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number or string|Die ID des PageContent-Objekts oder die Indexposition des PageContent-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[PageContent](pagecontent.md)

### getItemAt(index: number)
Ruft einen Seiteninhalt anhand seiner Position in der Auflistung ab.

#### Syntax
```js
pageContentCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[PageContent](pagecontent.md)

#### Beispiele
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;
    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("The first page content item is of type: " + firstPageContent.type);
            return context.sync();
        });
})
.catch(function(error) {
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

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
            });
        });
})                
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**Durchlaufen für Gliederungen**
```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
       .then(function () {    
              $.each(pageContents.items, function (index, pageContent) {
                     console.log(pageContent.type);
                     if (pageContent.type === 'Outline') {
                           outlines.push(pageContent);
                     }
              });
              $.each(outlines, function (index, outline) {
                     outline.load("id,paragraphs,paragraphs/type");
              });
              return context.sync();
       })
       .then(function () {
              $.each(outlines, function (index, outline) {
                     console.log("An outline was found with id : " + outline.id);
              });
              return Promise.resolve(outlines);
       });
});
```

