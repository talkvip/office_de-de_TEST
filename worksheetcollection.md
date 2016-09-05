# WorksheetCollection-Objekt (JavaScript-API für Excel)

Stellt eine Auflistung der Arbeitsblattobjekte dar, die Teil der Arbeitsmappe sind.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Items|[Worksheet[]](worksheet.md)|Eine Auflistung von Arbeitsblattobjekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Arbeitsblatt](worksheet.md)|Fügt der Arbeitsmappe ein neues Arbeitsblatt hinzu. Das Arbeitsblatt wird automatisch am Ende der vorhandenen Arbeitsblättern hinzugefügt. Wenn Sie das neu hinzugefügte Arbeitsblatt aktivieren möchten, rufen Sie ".activate() dazu auf.|
|[getActiveWorksheet()](#getactiveworksheet)|[Arbeitsblatt](worksheet.md)|Ruft das derzeit aktive Arbeitsblatt in der Arbeitsmappe ab.|
|[getItem(key: string)](#getitemkey-string)|[Arbeitsblatt](worksheet.md)|Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### add(name: string)
Fügt der Arbeitsmappe ein neues Arbeitsblatt hinzu. Das Arbeitsblatt wird automatisch am Ende der vorhandenen Arbeitsblättern hinzugefügt. Wenn Sie das neu hinzugefügte Arbeitsblatt aktivieren möchten, rufen Sie ".activate() dazu auf.

#### Syntax
```js
worksheetCollectionObject.add(name);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|name|string|Optional. Der Name des hinzuzufügenden Arbeitsblatts. Falls angegeben, sollte der Name eindeutig sein. Falls nicht angegeben, bestimmt Excel den Namen des neuen Arbeitsblatts.|

#### Gibt Folgendes zurück:
[Arbeitsblatt](worksheet.md)

#### Beispiele

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getActiveWorksheet()
Ruft das derzeit aktive Arbeitsblatt in der Arbeitsmappe ab.

#### Syntax
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### Parameter
Keine

#### Gibt Folgendes zurück:
[Arbeitsblatt](worksheet.md)

#### Beispiele

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(key: string)
Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab.

#### Syntax
```js
worksheetCollectionObject.getItem(key);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Key|string|Der Name oder die ID des Arbeitsblatts.|

#### Gibt Folgendes zurück:
[Arbeitsblatt](worksheet.md)

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
  var worksheets = ctx.workbook.worksheets;
  worksheets.load({"items" : "id, name"});
  return ctx.sync().then(function() {
    for (var i = 0; i < worksheets.items.length; i++)
    {
      console.log(worksheets.items[i].name);
      console.log(worksheets.items[i].id);
    }
  });
}).catch(function(error) {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```
