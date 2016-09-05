# ChartDataLabelFormat-Objekt (JavaScript-API für Excel)

Kapselt die Formateigenschaften für die Diagrammdatenbeschriftungen.

## Eigenschaften

Keine

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|fill|[ChartFill](chartfill.md)|Stellt die Füllung der aktuellen Diagrammdatenbeschriftung dar. Schreibgeschützt.|
|font|[ChartFont](chartfont.md)|Stellt die Zeichenformatierung (Schriftart, Schriftgrad, Farbe usw.) für eine Diagrammdatenbeschriftung dar. Schreibgeschützt.|

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
