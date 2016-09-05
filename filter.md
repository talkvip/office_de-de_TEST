# Filter-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Verwaltet das Filtern der Spalte einer Tabelle.

## Eigenschaften

Keine

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Kriterien|[FilterCriteria](filtercriteria.md)|Der aktuell angewendete Filter in der angegebenen Spalte. Schreibgeschützt.|

## Methods

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Wendet die angegebenen Filterkriterien in der angegebenen Spalte an. Die gleiche Funktionalität kann mit einer der folgenden Helper-Methoden erreicht werden.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Wendet den Filter "Bottom Item" auf die Spalte für die angegebene Anzahl von Elementen an.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Wendet den Filter "Bottom Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Wendet den Filter "Cell Color" auf die Spalte für die angegebene Farbe an.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Wendet den Filter "Symbol" auf die Spalte für die angegebenen Kriterienzeichenfolgen an.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Wendet den Filter "Dynamisch" auf die Spalte an.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Wendet den Filter "Font Color" auf die Spalte für die angegebene Farbe an.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Wendet den Filter "Icon" auf die Spalte für das angegebene Symbol an.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Wendet den Filter "Top Item" auf die Spalte für die angegebene Anzahl von Elementen an.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Wendet den Filter "Top Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues)|void|Wendet den Filter "Values" auf die Spalte für die angegebenen Werte an.|
|[clear()](#clear)|void|Deaktiviert den Filter für die angegebene Spalte.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### apply(criteria: FilterCriteria)
Wendet die angegebenen Filterkriterien in der angegebenen Spalte an. Die gleiche Funktionalität kann mit einer der folgenden Helper-Methoden erreicht werden. 

#### Syntax
```js
filterObject.apply(criteria);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Kriterien|FilterCriteria|Die Kriterien, die angewendet werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
Im folgenden Beispiel wird gezeigt, wie ein benutzerdefinierter Filter mit der generischen apply()-Methode angewendet wird.

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
        filterOn: Excel.FilterOn.custom,
        criterion1: ">50",
        operator: Excel.FilterOperator.and,
        criterion2: "<100"
        } 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
Wendet den Filter "Bottom Item" auf die Spalte für die angegebene Anzahl von Elementen an.

#### Syntax
```js
filterObject.applyBottomItemsFilter(count);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|count|number|Die Anzahl der Elemente vom unteren Rand, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
Wendet den Filter "Bottom Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.

#### Syntax
```js
filterObject.applyBottomPercentFilter(percent);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|percent|number|Der Prozentsatz von Elementen vom unteren Rand, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
Wendet den Filter "Cell Color" auf die Spalte für die angegebene Farbe an.


#### Syntax
```js
filterObject.applyCellColorFilter(color);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|color|string|Die Hintergrundfarbe der Zellen, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Wendet den Filter "Symbol" auf die Spalte für die angegebenen Kriterienzeichenfolgen an.

#### Syntax
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|criteria1|string|Die erste Kriterienzeichenfolge.|
|criteria2|string|Optional. Die zweite Kriterienzeichenfolge.|
|oper|FilterOperator|Optional. Der Operator, der beschreibt, wie die beiden Kriterien miteinander verknüpft sind.|

#### Returns
void


#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
Wendet den Filter "Dynamisch" auf die Spalte an.

#### Syntax
```js
filterObject.applyDynamicFilter(criteria);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Kriterien|string|Die dynamischen Kriterien, die angewendet werden sollen.  Die folgenden Werte sind möglich: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### Returns
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
Wendet den Filter "Font Color" auf die Spalte für die angegebene Farbe an.

#### Syntax
```js
filterObject.applyFontColorFilter(color);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|color|string|Die Schriftfarbe der Zellen, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon: Icon)
Wendet den Filter "Icon" auf die Spalte für das angegebene Symbol an.

#### Syntax
```js
filterObject.applyIconFilter(icon);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Symbol|Symbol|Die Symbole der Zellen, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
Wendet den Filter "Top Item" auf die Spalte für die angegebene Anzahl von Elementen an.

#### Syntax
```js
filterObject.applyTopItemsFilter(count);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|count|number|Die Anzahl der Elemente vom oberen Rand, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
Wendet den Filter "Top Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.

#### Syntax
```js
filterObject.applyTopPercentFilter(percent);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|percent|number|Der Prozentsatz von Elementen vom oberen Rand, die angezeigt werden sollen.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
Wendet den Filter "Values" auf die Spalte für die angegebenen Werte an.

#### Syntax
```js
filterObject.applyValuesFilter(values);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Werte|()[]|Die Liste der anzuzeigenden Werte.|

#### Rückgabewerte 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Deaktiviert den Filter für die angegebene Spalte.

#### Syntax
```js
filterObject.clear();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiel
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
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
