# FilterCriteria-Objekt (JavaScript-API für Excel)

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

Stellt die auf eine Spalte angewendeten Filterkriterien dar.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|color|string|Die HTML-Farbzeichenfolge, die zum Filtern von Zellen verwendet wird. Wird mit "cellColor"- und "fontColor"-Filterung verwendet.|
|criterion1|string|Das erste verwendete Kriterium zum Filtern von Daten. Wird als Operator bei "custom"-Filtern verwendet.|
|criterion2|string|Das zweite verwendete Kriterium zum Filtern von Daten. Wird nur als Operator bei "custom"-Filtern verwendet.|
|dynamicCriteria|string|Die dynamischen Kriterien aus dem Satz "Excel.DynamicFilterCriteria", die auf diese Spalte angewendet werden. Wird mit "dynamic"-Filtern verwendet. Die folgenden Werte sind möglich: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|
|filterOn|string|Die Eigenschaft, die vom Filter verwendet wird, um zu bestimmen, ob die Werte sichtbar bleiben sollen. Die folgenden Werte sind möglich:   BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom |
|values|object[]|Der Satz von Werten, der als Teil von "values"-Filtern verwendet werden soll.|

## Beziehungen
| Beziehung | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Symbol|[Symbol](icon.md)|Das zum Filtern von Zellen verwendete Symbol. Wird mit "icon"-Filtern verwendet.|
|operator|[FilterOperator](filteroperator.md)|Der Operator, mit dem Kriterium 1 und 2 bei Verwendung von "Custom"-Filtern kombiniert werden.|

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
