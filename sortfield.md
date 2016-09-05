# SortField-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt eine Bedingung in einem Sortiervorgang dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|ascending|bool|Stellt dar, ob die Sortierung in aufsteigender Reihenfolge ausgeführt wird.|
|color|string|Stellt die Farbe dar, die das Ziel der Bedingung ist, wenn die Sortierung für die Schrift- oder Zellenfarbe gilt.|
|dataOption|string|Stellt weitere Sortieroptionen für dieses Feld dar. Die folgenden Werte sind möglich: Normal, TextAsNumber.|
|key|int|Stellt die Spalte (oder Zeile, je nach Sortierausrichtung) dar, für die die Bedingung gilt. Wird als Offset von der ersten Spalte (oder Zeile) dargestellt.|
|sortOn|string|Stellt den Typ der Sortierung dieser Bedingung dar. Die folgenden Werte sind möglich: Value, CellColor, FontColor, Icon.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Symbol|[Symbol](icon.md)|Stellt das Symbol dar, das das Ziel der Bedingung ist, wenn die Sortierung für das Symbol der Zelle gilt.|

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
