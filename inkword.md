# InkWord-Objekt (JavaScript-API für OneNote)

_Betrifft: OneNote Online_  


Ein Container für die Freihand in einem Wort in einem Absatz.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Ruft die ID des InkWord-Objekts ab. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-id)|
|languageId|string|Die ID der erkannten Sprache in dieses Freihandwort. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-languageId)|
|wordAlternates|string|Die Wörter, die in diesem Freihandwort erkannt wurden, nach Wahrscheinlichkeit sortiert. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-wordAlternates)|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Der übergeordnete Absatz, der das Freihandwort enthält. Schreibgeschützt.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-paragraph)|

## Methoden

| Methode           | Rückgabetyp    |Beschreibung| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|[OK](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWord-load)|

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
