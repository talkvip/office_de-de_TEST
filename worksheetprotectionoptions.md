# WorksheetProtectionOptions-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt die Optionen für den Arbeitsblattschutz dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|allowAutoFilter|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung der automatischen Filterfunktion dar.|
|allowDeleteColumns|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Löschens von Spalten dar.|
|allowDeleteRows|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Löschens von Zeilen dar.|
|allowFormatCells|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Zellen dar.|
|allowFormatColumns|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Spalten dar.|
|allowFormatRows|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Zeilen dar.|
|allowInsertColumns|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Spalten dar.|
|allowInsertHyperlinks|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Hyperlinks dar.|
|allowInsertRows|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Zeilen dar.|
|allowPivotTables|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung der Pivottabellenfunktion dar.|
|allowSort|bool|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung der Sortierfunktion dar.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

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
In diesem Beispiel werden die Schutzoptionen für das aktive Arbeitsblatt geladen.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
