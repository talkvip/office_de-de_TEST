# ParagraphCollection-Objekt (JavaScript-API für Word)

Enthält eine Sammlung von [Absatz](paragraph.md)Objekten.

__Gilt für: Word 2016, Word für iPad, Word für Mac_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Items|[Paragraph[]](paragraph.md)|Eine Sammlung von Abschnitt-Objekten. Schreibgeschützt.|

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

#### Beispiele
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text and style properties for all of the paragraphs.
    context.load(paragraphs, 'text, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the last paragraph and create a 
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1]; 
        
        // Queue a command to select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Supportdetails
Verwenden Sie den [Anforderungssatz](../office-add-in-requirement-sets.md) in Laufzeitüberprüfungen, um sicherzustellen, dass die Anwendung von der Hostversion von Word unterstützt wird. Weitere Informationen zur Office-Hostanwendung und den Serveranforderungen finden Sie unter [Anforderungen für die Ausführung von Office-Add-Ins](../../docs/overview/requirements-for-running-office-add-ins.md).