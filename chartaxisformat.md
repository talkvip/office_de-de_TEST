# ChartAxisFormat-Objekt (JavaScript-API für Excel)

Kapselt die Formateigenschaften für die Diagrammachse.

## Eigenschaften

Keine

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|font|[ChartFont](chartfont.md)|Stellt die Zeichenformatierung wie Schriftart, Schriftgrad, Farbe usw. des Diagrammachsenelements dar. Schreibgeschützt.|
|line|[ChartLineFormat](chartlineformat.md)|Stellt die Formatierung der Diagrammlinien dar. Schreibgeschützt.|

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
