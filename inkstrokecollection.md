# InkStrokeCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_   


Stellt eine Auflistung von InkStroke-Objekten dar.


## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl von InkStrokes auf der Seite zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-count)|
|Items|[InkStroke[]](inkstroke.md)|Eine Auflistung von inkStroke-Objekten.
 Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[InkStroke](inkstroke.md)|Ruft ein InkStroke-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkStroke](inkstroke.md)|Ruft einen InkStroke anhand seiner Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-load)|

## Methodendetails


### getItem(index: number oder string)
Ruft ein InkStroke-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
inkStrokeCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number oder string|Die ID des InkStroke-Objekts oder die Indexposition des InkStroke-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[InkStroke](inkstroke.md)

### getItemAt(index: number)
Ruft einen InkStroke anhand seiner Position in der Auflistung ab.

#### Syntax
```js
inkStrokeCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[InkStroke](inkstroke.md)

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
