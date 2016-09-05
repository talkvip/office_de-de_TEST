# RequestContext-Objekt (JavaScript-API für Excel)

Das RequestContext-Objekt erleichtert das Senden von Anforderungen an die Excel-Anwendung. Da das Office-Add-In und die Excel-Anwendung in zwei verschiedenen Prozessen ausgeführt werden, ist Anforderungskontext erforderlich, um auf Excel und die zugehörigen Objekte, wie z. B. Arbeitsblätter, Tabellen usw. aus dem Add-In zuzugreifen. 

## Eigenschaften
Keine

## Methoden

| Methode         | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Optionen.|

## API-Spezifikation

### load(object: object, option: object)
Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Optionen.

#### Syntax
```js
requestContextObject.load(object, loadOption);
```

#### Parameter
| Parameter       | Typ    |Beschreibung|
|:----------------|:--------|:----------|
|object|object|Optional. Geben Sie den Namen des zu ladenden Objekts an.|
|Option|[loadOption](loadoption.md)|Optional. Geben Sie die Ladeoptionen wie z. B. das Auswählen, Erweitern, Überspringen und nach oben an. Weitere Informationen finden Sie unter dem loadOption-Objekt.|

#### Gibt 
void

##### Beispiele

Im folgenden Beispiel werden Eigenschaftswerte aus einem Bereich geladen und in einen anderen Bereich kopiert.

```js
Excel.run(function (ctx) { 
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
    ctx.load(range, "values");
    return ctx.sync().then(function() {
        var myvalues=range.values;
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
        console.log(range.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
