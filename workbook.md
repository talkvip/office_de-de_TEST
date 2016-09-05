# Arbeitsmappenobjekt (JavaScript-API für Excel)

Die Arbeitsmappe ist das Objekt auf oberster Ebene , das dazugehörige Arbeitsmappenobjekte wie z. B. Arbeitsblätter, Tabellen, Bereiche usw. enthält.

## Eigenschaften

Keine

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|application|[Anwendung](application.md)|Stellt eine Instanz der Excel-Anwendung dar, die diese Arbeitsmappe enthält. Schreibgeschützt.|
|bindings|[BindingCollection](bindingcollection.md)|Stellt eine Auflistung aller Bindungen dar, die Teil der Arbeitsmappe sind. Schreibgeschützt.|
|functions|[Funktionen](functions.md)|Stellt die Instanz der Excel-Anwendung dar, die diese Arbeitsmappe enthält. Schreibgeschützt.|
|names|[NamedItemCollection](nameditemcollection.md)|Stellt eine Auflistung der benannten Elemente des Arbeitsmappenbereichs dar (benannte Bereiche und Konstanten). Schreibgeschützt.|
|Tabellen|[TableCollection](tablecollection.md)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Tabellen dar. Schreibgeschützt.|
|Worksheets|[WorksheetCollection](worksheetcollection.md)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Arbeitsblätter dar. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Bereichen](range.md)|Ruft den aktuell ausgewählten Bereich aus der Arbeitsmappe ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getSelectedRange()
Ruft den aktuell ausgewählten Bereich aus der Arbeitsmappe ab.

#### Syntax
```js
workbookObject.getSelectedRange();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Bereichen](range.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
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
