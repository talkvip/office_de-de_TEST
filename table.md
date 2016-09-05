# Table object (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt eine Excel-Tabelle dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|id|int|Gibt einen Wert zurück, der das Arbeitsblatt in einer bestimmten Arbeitsmappe eindeutig identifiziert. Der Wert des Bezeichners bleibt unverändert, auch wenn die Tabelle umbenannt wird. Schreibgeschützt.|
|name|string|Der Name der Tabelle.|
|showHeaders|bool|Gibt an, ob die Kopfzeile sichtbar oder nicht sichtbar ist. Dieser Wert kann festgelegt werden, um die Kopfzeile anzuzeigen, oder sie zu entfernen.|
|showTotals|bool|Gibt an, ob die Ergebniszeile sichtbar ist oder nicht. Dieser Wert kann so festgelegt werden, dass die Ergebniszeile angezeigt oder ausgeblendet wird.|
|style|string|Konstanter Wert, der das Tabellenformat darstellt. Die folgenden Werte sind möglich: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. Es kann ebenfalls eine in der Arbeitsmappe vorhandene benutzerdefinierte Formatierung angegeben werden.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|Stellt eine Auflistung aller Spalten in der Tabelle dar. Schreibgeschützt.|
|Rows|[TableRowCollection](tablerowcollection.md)|Stellt eine Auflistung aller Zeilen in der Tabelle dar. Schreibgeschützt.|
|sort|[TableSort](tablesort.md)|Stellt die Sortierkonfiguration für die Tabelle dar. Schreibgeschützt.|
|worksheet|[Arbeitsblatt](worksheet.md)|Das Arbeitsblatt, das die aktuelle Tabelle enthält. Schreibgeschützt.|

## Methods

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|Löscht alle Filter, die derzeit in der Tabelle verwendet werden.|
|[convertToRange()](#converttorange)|[Bereichen](range.md)|Wandelt die Tabelle in einen normalen Bereich von Zellen um. Alle Daten werden beibehalten.|
|[delete()](#delete)|void|Löscht die Tabelle.|
|[getDataBodyRange()](#getdatabodyrange)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das mit dem Datenteil der Tabelle verknüpft ist.|
|[getHeaderRowRange()](#getheaderrowrange)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das mit der Kopfzeile der Tabelle verknüpft ist.|
|[getRange()](#getrange)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das mit der gesamten Tabelle verknüpft ist.|
|[getTotalRowRange()](#gettotalrowrange)|[Bereichen](range.md)|Ruft das Bereichsobjekt ab, das mit der Ergebniszeile der Tabelle verknüpft ist.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|
|[reapplyFilters()](#reapplyfilters)|void|Wendet alle Filter erneut an, die derzeit in der Tabelle vorhanden sind.|

## Methodendetails


### clearFilters()
Löscht alle Filter, die derzeit in der Tabelle verwendet werden.

#### Syntax
```js
tableObject.clearFilters();
```

#### Parameter
Keine

#### Gibt 
void

### convertToRange()
Wandelt die Tabelle in einen normalen Bereich von Zellen um. Alle Daten werden beibehalten.

#### Syntax
```js
tableObject.convertToRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### delete()
Löscht die Tabelle.

#### Syntax
```js
tableObject.delete();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getDataBodyRange()
Ruft das Bereichsobjekt ab, das mit dem Datenteil der Tabelle verknüpft ist.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getHeaderRowRange()
Ruft das Bereichsobjekt ab, das mit der Überschriftenzeile der Tabelle verknüpft ist.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange()
Ruft das Bereichsobjekt ab, das mit der gesamten Tabelle verknüpft ist.

#### Syntax
```js
tableObject.getRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address'); 
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTotalRowRange()
Ruft das Bereichsobjekt ab, das mit der Ergebniszeile der Tabelle verknüpft ist.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');   
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
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

Rufen Sie eine Tabelle nach Namen ab. 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Rufen Sie eine Tabelle nach Index ab.

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.name('name')
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

Legen Sie ein Tabellenformat fest 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.tableStyle = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.tableStyle);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
