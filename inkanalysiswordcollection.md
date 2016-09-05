# InkAnalysisWordCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt eine Auflistung von InkAnalysisWord-Objekten dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl von InkAnalysisWords auf der Seite zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-count)|
|Items|[InkAnalysisWord[]](inkanalysisword.md)|Eine Auflistung von inkAnalysisWord-Objekten. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[InkAnalysisWord](inkanalysisword.md)|Ruft ein InkAnalysisWord-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisWord](inkanalysisword.md)|Ruft ein InkAnalysisWord anhand seiner Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-load)|

## Methodendetails


### getItem(index: number oder string)
Ruft ein InkAnalysisWord-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
inkAnalysisWordCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number oder string|Die ID des InkAnalysisWord-Objekts oder die Indexposition des InkAnalysisWord-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[InkAnalysisWord](inkanalysisword.md)

### getItemAt(index: number)
Ruft ein InkAnalysisWord anhand seiner Position in der Auflistung ab.

#### Syntax
```js
inkAnalysisWordCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[InkAnalysisWord](inkanalysisword.md)

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
