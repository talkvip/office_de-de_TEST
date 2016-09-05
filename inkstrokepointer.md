# InkStrokePointer-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Weak-Verweise auf ein Freihandstrichobjekt und sein übergeordnetes Inhaltselement.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|contentId|string|Stellt die ID des Seiteninhaltsobjekts dar, das diesem Strich entspricht.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|Stellt die ID des Freihandstrichs dar.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

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
