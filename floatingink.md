# FloatingInk-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt eine Gruppe von Freihandstrichen dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Ruft die ID des FloatingInk-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|Ruft die Striche des FloatingInk-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[PageContent](pagecontent.md)|Ruft das übergeordnete Element PageContent des FloatingInk-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

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

**id**
```js
OneNote.run(function(context) {

    // Gets the active page.
    var page = context.application.getActivePage();
    var contents = page.contents;
    
    // Load page contents and their types.
    page.load('contents/type');
    return context.sync()
        .then(function(){
        
            // Load every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    content.load('ink/id');
                }                           
            })
            return context.sync();
        })
        .then(function(){
        
            // Log ID of every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    console.log(content.ink.id);
                }                           
            })              
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
