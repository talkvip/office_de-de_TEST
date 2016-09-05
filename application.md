# Application-Objekt (JavaScript-API für Excel)

Stellt die Excel-Anwendung dar, die die Arbeitsmappe verwaltet.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|calculationMode|string|Gibt den in der Arbeitsmappe verwendeten Berechnungsmodus zurück. Schreibgeschützt. Die folgenden Werte sind möglich: `Automatic`: Excel steuert Neuberechnung, `AutomaticExceptTables`: Excel steuert Neuberechnung, ignoriert jedoch die Änderungen in Tabellen, `Manual`: Berechnung wird auf Benutzeranforderung durchgeführt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Alle in Excel geöffnete Arbeitsmappen werden neu berechnet.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### calculate(calculationType: string)
Alle in Excel geöffnete Arbeitsmappen werden neu berechnet.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|calculationType|string|Gibt den zu verwendenden Berechnungstyp an. Die folgenden Werte sind möglich: `Recalculate` Standardoption: Führt eine normale Berechnung durch, indem alle Formeln in der Arbeitsmappe berechnet werden. `Full`: Erzwingt eine vollständige Datenberechnung. `FullRebuild`: Erzwingt eine vollständige Datenberechnung und aktualisiert die Abhängigkeiten.|

#### Gibt 
void

#### Beispiele
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
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
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder akzeptiert ein [loadOption](loadoption.md)-Objekt.|

#### Gibt 
void
### Eigenschaftszugriffsbeispiele
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

