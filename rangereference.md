# RangeReference-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt einen Zeichenfolgenverweis der Form Arbeitsblattname!A1:B5 oder einen globalen oder lokalen benannten Bereich dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|address|string|Das Arbeitsblatt, das den aktuellen Bereich enthält.|

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
