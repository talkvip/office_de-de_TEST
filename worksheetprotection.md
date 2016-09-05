# WorksheetProtection-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Steht den Schutz eines Arbeitsblattobjekts dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|geschützt|bool|Zeigt an, ob das Arbeitsblatt geschützt ist. Schreibgeschützt.|

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Optionen|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Optionen für den Arbeitsblattschutz. Schreibgeschützt.|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Füllt das Proxyobjekt mit Details zum Schutz des Blatts.|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|Schützen ein Arbeitsblatt. Wird ausgelöst, wenn das Arbeitsblatt geschützt ist.|
|[unprotect()](#unprotect)|void|Schutz eines Arbeitsblatts aufheben.|

## Methodendetails


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

#### Beispiele
In diesem Beispiel werden die Schutzdetails für das aktive Arbeitsblatt geladen.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### protect(options: WorksheetProtectionOptions)
Schützt ein Arbeitsblatt mit optionalen Schutzrichtlinien. Eine Ausnahme wird ausgelöst, wenn das Arbeitsblatt geschützt ist. 

Wenn Optionen angegeben werden, können einzelne Richtlinien aktiviert oder deaktiviert werden. Wenn eine Richtlinie nicht angegeben ist, ist sie standardmäßig aktiviert. 

#### Syntax
```js
worksheetProtectionObject.protect(options);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Optionen|WorksheetProtectionOptions|Optionale Optionen für den Arbeitsblattschutz.|


#### Rückgabewerte 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### unprotect()
Schutz eines Arbeitsblatts aufheben. 

#### Syntax
```js
worksheetProtectionObject.unprotect();
```

#### Parameter
Keine

#### Gibt 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");  
    sheet.protection.unprotect();
    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```