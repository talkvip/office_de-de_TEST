# TableRow-Objekt (JavaScript-API für Excel)

Stellt eine Zeile in einer Tabelle dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Index|int|Gibt die Indexnummer der Zeile in der Zeilenauflistung der Tabelle zurück. Nullindiziert. Schreibgeschützt.|
|values|object[][]|Stellt die Rohwerte des angegebenen Bereichs dar. Die zurückgegebenen Daten können den Typ Zeichenfolge, Zahl oder ein boolescher Wert sein. Eine Zelle, die einen Fehler enthält, gibt eine Fehlerzeichenfolge zurück.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Löscht die Zeile aus der Tabelle.|
|[getRange()](#getrange)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das mit der gesamten Zeile verknüpft ist.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### delete()
Löscht die Zeile aus der Tabelle.

#### Syntax
```js
tableRowObject.delete();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
    row.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange()
Ruft das Bereichsobjekt ab, das mit der gesamten Zeile verknüpft ist.

#### Syntax
```js
tableRowObject.getRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
    var rowRange = row.getRange();
    rowRange.load('address');
    return ctx.sync().then(function() {
        console.log(rowRange.address);
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
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItem(0);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New", "Values", "For", "New", "Row"]];
    var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
    row.values = newValues;
    row.load('values');
    return ctx.sync().then(function() {
        console.log(row.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
