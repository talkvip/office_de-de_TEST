# FormatProtection-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt den Formatschutz eines Bereichsobjekts dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|formulaHidden|bool|Zeigt an, ob Excel die Formel für die Zellen im Bereich ausblendet. Ein Nullwert gibt an, dass der gesamte Bereich keine einheitliche Einstellung zur Formelausblendung hat.|
|locked|bool|Gibt an, ob Excel die Zellen im Objekt sperrt. Ein Nullwert gibt an, dass der gesamte Bereich keine einheitliche Einstellung zum Sperren hat.|

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
