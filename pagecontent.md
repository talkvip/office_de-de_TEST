# PageContent-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt einen Bereich auf einer Seite dar, der Inhaltstypen auf oberster Ebene enthält, z. B. Outline oder Image. Ein PageContent-Objekt kann einer XY-Position zugewiesen werden.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Ruft die ID des PageContent-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|Ruft die linke Position (X-Achse) des PageContent-Objekts ab oder legt sie fest.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|Ruft die obere Position (X-Achse) des PageContent-Objekts ab oder legt sie fest.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|Ruft den Typ des PageContent-Objekts ab. Schreibgeschützt. Die folgenden Werte sind möglich: Outline, Image, Other.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|image|[Bild](image.md)|Ruft das Bild des PageContent-Objekts ab. Löst eine Ausnahme aus, wenn PageContentType nicht vom Typ „Image“ ist. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|Freihand|[FloatingInk](floatingink.md)|Ruft die Freihand im PageContent-Objekt ab. Löst eine Ausnahme aus, wenn PageContentType nicht vom Typ „Ink“ ist. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Gliederung](outline.md)|Ruft die Gliederung des PageContent-Objekts ab. Löst eine Ausnahme aus, wenn PageContentType nicht vom Typ „Outline“ ist. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|Ruft die Seite ab, die das PageContent-Objekt enthält. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Löscht das PageContent-Objekt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## Methodendetails


### delete()
Löscht das PageContent-Objekt.

#### Syntax
```js
pageContentObject.delete();
```

#### Parameter
Keine

#### Gibt 
void

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
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
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
