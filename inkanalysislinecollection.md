# InkAnalysisLineCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt eine Auflistung von InkAnalysisLine-Objekten dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl von InkAnalysisLines auf der Seite zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|Items|[InkAnalysisLine[]](inkanalysisline.md)|Stellt eine Auflistung von inkAnalysisLine-Objekten dar. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[InkAnalysisLine](inkanalysisline.md)|Ruft ein InkAnalysisLine-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|Ruft eine InkAnalysisLine anhand ihrer Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## Methodendetails


### getItem(index: number oder string)
Ruft ein InkAnalysisLine-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number oder string|Die ID des InkAnalysisLine-Objekts oder die Indexposition des InkAnalysisLine-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[InkAnalysisLine](inkanalysisline.md)

### getItemAt(index: number)
Ruft eine InkAnalysisLine anhand ihrer Position in der Auflistung ab.

#### Syntax
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[InkAnalysisLine](inkanalysisline.md)

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
