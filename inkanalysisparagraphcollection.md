# InkAnalysisParagraphCollection-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Stellt eine Auflistung von InkAnalysisParagraph-Objekten dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Gibt die Anzahl von InkAnalysisParagraphs auf der Seite zurück. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-count)|
|Items|[InkAnalysisParagraph[]](inkanalysisparagraph.md)|Eine Auflistung von InkAnalysisParagraph-Objekte. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-items)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number oder string)](#getitemindex-number-oder-string)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Ruft ein InkAnalysisParagraph-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Ruft einen InkAnalysisParagraph anhand seiner Position in der Auflistung ab.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-load)|

## Methodendetails


### getItem(index: number oder string)
Ruft ein InkAnalysisParagraph-Objekt nach ID oder dem Index in der Auflistung ab. Schreibgeschützt.

#### Syntax
```js
inkAnalysisParagraphCollectionObject.getItem(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number oder string|Die ID des InkAnalysisParagraph-Objekts oder die Indexposition des InkAnalysisParagraph-Objekts in der Auflistung.|

#### Gibt Folgendes zurück:
[InkAnalysisParagraph](inkanalysisparagraph.md)

### getItemAt(index: number)
Ruft einen InkAnalysisParagraph anhand seiner Position in der Auflistung ab.

#### Syntax
```js
inkAnalysisParagraphCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[InkAnalysisParagraph](inkanalysisparagraph.md)

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
