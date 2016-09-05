# NamedItemCollection-Objekt (JavaScript-API für Excel)

Eine Auflistung aller nameditem-Objekte, die Teil der Arbeitsmappe sind.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Items|[NamedItem[]](nameditem.md)|Eine Auflistung von namedItem-Objekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Ruft eine nameditem-Objekt, das den Namen verwendet|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getItem(name: string)
Ruft eine nameditem-Objekt, das den Namen verwendet

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|name|string|nameditem-Name.|

#### Gibt Folgendes zurück:
[NamedItem](nameditem.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var nameditem = ctx.workbook.names.getItem(wSheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

#### Beispiele

```js
Excel.run(function (ctx) { 
    var nameditem = ctx.workbook.names.getItemAt(0);
    nameditem.load('name');
    return ctx.sync().then(function() {
            console.log(nameditem.name);
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
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder geben Sie das [loadOption](loadoption.md)-Objekt an.|

#### Gibt 
void
### Eigenschaftszugriffsbeispiele

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Rufen Sie die Anzahl der benannten Elemente ab.

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('count');
    return ctx.sync().then(function() {
        console.log("nameditems: Count= " + nameditems.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

