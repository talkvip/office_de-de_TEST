# ChartAxes-Objekt (JavaScript-API für Excel)

Die Achsen des Diagramms.

## Eigenschaften

Keine

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|Stellt die Rubrikenachse in einem Diagramm dar. Schreibgeschützt.|
|seriesAxis|[ChartAxis](chartaxis.md)|Stellt die Reihenachse eines dreidimensionalen Diagramms dar. Schreibgeschützt.|
|valueAxis|[ChartAxis](chartaxis.md)|Stellt die Größenachse in einer Achse dar. Schreibgeschützt.|

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