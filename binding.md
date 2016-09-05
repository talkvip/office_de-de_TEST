# Binding-Objekt (JavaScript-API für Excel)

Stellt eine Office.js-Bindung dar, die in der Arbeitsmappe definiert wird.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|id|string|Stellt die Bindungs-ID dar. Schreibgeschützt.|
|type|string|Gibt den Typ der Bindung an. Schreibgeschützt. Die folgenden Werte sind möglich: Bereich, Tabelle, Text.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Bereichen](range.md)|Gibt den durch die Bindung dargestellten Bereich zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.|
|[getTable()](#gettable)|[Tabelle](table.md)|Gibt die durch die Bindung dargestellte Tabelle zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.|
|[getText()](#gettext)|string|Gibt den durch die Bindung dargestellten Text zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getRange()
Gibt den durch die Bindung dargestellten Bereich zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.

#### Syntax
```js
bindingObject.getRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
Im folgenden Beispiel wird ein Binding-Objekt zum Abrufen des verknüpften Bereichs verwendet.

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTable()
Gibt die durch die Bindung dargestellte Tabelle zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.

#### Syntax
```js
bindingObject.getTable();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Tabelle](table.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getText()
Gibt den durch die Bindung dargestellten Text zurück. Gibt einen Fehler zurück, wenn die Bindung nicht den korrekten Typ aufweist.

#### Syntax
```js
bindingObject.getText();
```

#### Parameter
Keine

#### Gibt 
string

#### Beispiele

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    ctx.load('text');
    return ctx.sync().then(function() {
        console.log(text);
    });
}).catch(function(error) {
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
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder akzeptiert ein [loadOption](loadoption.md)-Objekt.|

#### Gibt 
void
### Eigenschaftszugriffsbeispiele

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
    return ctx.sync().then(function() {
        console.log(binding.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
