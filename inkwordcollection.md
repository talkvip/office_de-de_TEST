# InkWordCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt eine Auflistung von InkWord-Objekten dar.


## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl von InkWords auf der Seite zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|Items|[InkWord[]](inkword.md)|Eine Auflistung von inkWord-Objekten.
 Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[InkWord](inkword.md)|Ruft ein InkWord-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|Ruft ein InkWord anhand seiner Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## Methodendetails


### getItem(index: number oder string)
Ruft ein InkWord-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
inkWordCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number oder string|Die ID des InkWord-Objekts oder die Indexposition des InkWord-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[InkWord](inkword.md)

### getItemAt(index: number)
Ruft ein InkWord anhand seiner Position in der Auflistung ab.

#### Syntax
```js
inkWordCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[InkWord](inkword.md)

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
