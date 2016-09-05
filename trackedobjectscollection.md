# TrackedObjectsCollection-Objekt (JavaScript-API für Office 2016)

Damit können Add-Ins die Bereichsobjektverweise über sync()-Batches hinweg verwalten. In der Regel können Sie mit der Excel.run() batchübergreifend Verweise automatisch beibehalten, ohne diese explizit nachzuverfolgen. Wenn ein Bereichsobjekt für ein Add-In-Szenario jedoch manuell nachverfolgt und angepasst werden muss, damit der aktuelle Status des zugrunde liegenden Excel-Bereichs widergespiegelt wird, kann diese Auflistung verwendet werden, um solche Objekte für die Überwachung zu markieren. Beachten Sie, dass ein Bereichsobjekt, das zur Überwachung markiert ist, bei der Verwendung explizit entfernt werden muss, um Speicherplatz in Excel freizugeben, selbst im Falle eines Fehlers.

## Eigenschaften
Keine.

## Beziehungen

Keine

## Methoden

Bei dem trackedObjectsCollection-Objekt sind folgende Methoden definiert:

| Methode     | Rückgabetyp    |Beschreibung|
|:-----------------|:--------|:----------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Erstellt einen neuen Verweis zu einem Bereich.|
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Entfernen Sie einen Verweis zu dem Bereich.  |
|[removeAll()](#removeall)| Null|Entfernt alle Verweise, die auf dem Gerät vom Add-In erstellt wurden.|


## API-Spezifikation 

### add(rangeObject: range)
Fügen Sie der trackedObjectsCollection ein Bereichsobjekt hinzu. Zugrunde liegende Änderungen an batchübergreifenden Anforderungen werden nachverfolgt und bei nachfolgenden Updates auf den aktuellen Status des Bereichsobjekts angewendet. 

#### Syntax
```js
trackedObjectsCollection.add(rangeObject);
```

#### Parameter

Parameter       | Typ   | Beschreibung
--------------- | ------ | ------------
`rangeObject`  | [Bereichen](range.md)| Das Bereichsobjekt, das zur trackedObjectCollection hinzugefügt werden muss.

#### Gibt 
Null

#### Beispiele

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    return ctx.sync(); 
});
```


### remove(rangeObject: range)

Entfernen Sie ein Verweisobjekt aus der Auflistung. Dadurch werden Speicherplatz und Ressourcen freigegeben, die zum Beibehalten des Status des nachverfolgten Objekts erforderlich sind. Beachten Sie, dass ein Bereichsobjekt, das zur Überwachung markiert ist, explizit entfernt werden muss, sogar im Falle eines Fehlers.

#### Syntax
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Parameter

Parameter       | Typ   | Beschreibung
--------------- | ------ | ------------
`rangeObject`  | [Bereichen](range.md)| Das Bereichsobjekt, das aus der trackedObjectCollection entfernt werden muss.

#### Gibt 
Null

#### Beispiele


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.remove(range); 
    return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

Entfernt alle Verweise, die auf dem Gerät vom Add-In erstellt wurden.

#### Syntax
```js
trackedObjectsCollection.removeAll();
```

#### Parameter

Keine

#### Gibt 
Null

#### Beispiele

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var ctx = new Excel.RequestContext();
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    ctx.trackedObjectsCollection.add(range);
    ctx.load(range);
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.removeAll(); 
    return ctx.sync(); 
});
```
