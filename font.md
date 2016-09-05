# Font-Objekt (JavaScript-API für Word)

Stellt eine Schriftart dar.

__Gilt für: Word 2016, Word für iPad, Word für Mac_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|bold|bool|Ruft einen Wert ab, der angibt, ob die Schriftart fett formatiert ist, oder legt diesen fest. True, wenn die Schriftart fett formatiert ist, andernfalls false.|
|color|string|Ruft die Farbe für die angegebene Schriftart ab, oder legt diese fest. Sie können den Wert im Format „#RRGGBB“ oder den Namen der Farbe angeben.|
|doubleStrikeThrough|bool|Ruft einen Wert ab, der angibt, ob die Schriftart doppelt durchgestrichen ist, oder legt diesen fest. True, wenn der Text doppelt durchgestrichen ist, andernfalls false.|
|highlightColor|string|Ruft die Markierungsfarbe für die angegebene Schriftart ab, oder legt diese fest. Sie können den Wert im Format „#RRGGBB“ oder den Namen der Farbe angeben.|
|italic|bool|Ruft einen Wert ab, der angibt, ob die Schriftart kursiv ist, oder legt diesen fest. True, wenn die Schriftart kursiv ist, andernfalls false.|
|name|string|Ruft einen Wert ab, der den Namen der Schriftart angibt, oder legt diesen fest.|
|strikeThrough|bool|Ruft einen Wert ab, der angibt, ob die Schriftart durchgestrichen ist, oder legt diesen fest. True, wenn der Text durchgestrichen ist, andernfalls false.|
|subscript|bool|Ruft einen Wert ab, der angibt, ob die Schriftart tiefgestellt ist, oder legt diesen fest. True, wenn die Schriftart tiefgestellt ist, andernfalls false.|
|superscript|bool|Ruft einen Wert ab, der angibt, ob die Schriftart hochgestellt ist, oder legt diesen fest. True, wenn die Schriftart hochgestellt ist, andernfalls false.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|size|**Gleitkommazahl**|Ruft einen Wert ab, der den Schriftgrad in Punkten angibt, oder legt diesen fest.|
|underline|**string**|Ruft einen Wert ab, der den auf die Schriftart angewendete Art der Unterstreichung angibt, oder legt diesen fest. Gültige Werte sind: „None“, „Single“, „Word“, „Double“, „Dotted“, „Hidden“, „Thick“, „Dashline“, „Dotline“, „DotDashLine“, „TwoDotDashLine“ und „Wave“|

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

    // Queue a commmand to load the font property for all of the paragraphs.
    context.load(paragraphs, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object for the font object on the first paragraph in the collection.
        var font = paragraphs.items[0].font;

        // Queue a set of property value changes on the font proxy object.
        font.size = 32;
        font.bold = true;
        font.color = '#0000ff';
        font.highlightColor = '#ffff00';

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('The font has changed.');
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

## Eigenschaftszugriffsbeispiele

### Ändern des Schriftartnamens
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font name.
    selection.font.name = 'Arial';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font name has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Ändern der Schriftfarbe
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Ändern des Schriftgrads
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font size.
    selection.font.size = 20;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font size has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Markieren von ausgewähltem Text
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection has been highlighted.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Fettformatieren von Text
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to make the current selection bold.
    selection.font.bold = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection is now bold.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### Unterstreichen von Text
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Durchstreichen von Text
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to strikethrough the font of the current selection.
    selection.font.strikeThrough = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has a strikethrough.');
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
