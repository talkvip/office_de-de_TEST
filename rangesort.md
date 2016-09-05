# RangeSort-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Verwaltet Sortiervorgänge für Range-Objekte.

## Eigenschaften

Keine

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Führt einen Sortiervorgang aus.|

## Methodendetails


### apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Führt einen Sortiervorgang aus.

#### Syntax
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|fields|SortField[]|Die Liste der Konditionen, nach denen sortiert werden soll.|
|matchCase|bool|Optional. Gibt an, ob sich die Groß-/Kleinschreibung auf die Zeichenfolgensortierung auswirkt.|
|hasHeaders|bool|Optional. Gibt an, ob der Bereich eine Kopfzeile aufweist.|
|orientation|string|Optional. Gibt an, ob der Vorgang Zeilen oder Spalten sortiert.  Die folgenden Werte sind möglich: Zeilen, Spalten|
|method|string|Optional. Die Sortiermethode für chinesische Zeichen.  Die folgenden Werte sind möglich: PinYin, StrokeCount|

#### Gibt Folgendes zurück:
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```