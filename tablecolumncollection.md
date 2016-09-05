# TableColumnCollection-Objekt (JavaScript-API für Excel)

Stellt eine Auflistung aller Spalten dar, die Teil der Tabelle sind.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|count|int|Gibt die Anzahl der Spalten in der Tabelle an. Schreibgeschützt.|
|Items|[TableColumn[]](tablecolumn.md)|Eine Auflistung von tableColumn-Objekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Fügt der Tabelle eine neue Spalte hinzu.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Ruft ein Spaltenobjekt nach Name oder ID ab.|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Ruft eine Spalte anhand ihrer Position in der Auflistung ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### add(index: number, values: (boolean or string or number)[][])
Fügt der Tabelle eine neue Spalte hinzu.

#### Syntax
```js
tableColumnCollectionObject.add(index, values);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Gibt die relative Position der neuen Spalte an. Die vorherige Spalte an dieser Position wird nach rechts verschoben. Der Indexwert sollte gleich oder kleiner als der letzten Spalten-Indexwert sein, damit er nicht dazu verwendet werden kann, eine Spalte an das Ende der Tabelle anzuhängen. Nullindiziert.|
|values|(boolean or string or number)[][]|Optional. Ein 2-dimensionales Array von unformatierten Werten der Tabellenspalte.|

#### Gibt Folgendes zurück:
[TableColumn](tablecolumn.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(key: number or string)
Ruft ein Spaltenobjekt nach Name oder ID ab.

#### Syntax
```js
tableColumnCollectionObject.getItem(key);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Key|Nummer oder Zeichenfolge| Spaltenname oder -ID.|

#### Gibt Folgendes zurück:
[TableColumn](tablecolumn.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getItemAt(index: number)
Ruft eine Spalte anhand ihrer Position in der Auflistung ab.

#### Syntax
```js
tableColumnCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[TableColumn](tablecolumn.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
