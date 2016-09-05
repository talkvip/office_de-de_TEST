# BindingCollection-Objekt (JavaScript-API für Excel)

Eine Sammlung aller Binding-Objekte, die Teil der Arbeitsmappe sind.

## Eigenschaften

| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|count|int|Gibt die Anzahl der Bindungen in der Sammlung zurück. Schreibgeschützt.|
|Items|[Binding[]](binding.md)|Eine Sammlung von Binding-Objekten. Schreibgeschützt.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#beispielen.)_

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Bindung](binding.md)|Ruft ein Binding-Objekt anhand seiner ID ab.|
|[getItemAt(index: number)](#getitematindex-number)|[Bindung](binding.md)|Ruft ein Binding-Objekt anhand seiner Position im Elementarray ab.|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails


### getItem(id: string)
Ruft ein Binding-Objekt anhand seiner ID ab.

#### Syntax
```js
bindingCollectionObject.getItem(id);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|id|string|ID des abzurufenden Binding-Objekts.|

#### Gibt Folgendes zurück:
[Bindung](binding.md)

#### Beispiele

Erstellen Sie eine Tabellenbindung für die Überwachung der Datenänderungen in der Tabelle. Wenn Daten geändert werden, ändert sich die Hintergrundfarbe der Tabelle zu Orange.

```js
(function () {
    // Create myTable
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.add("Sheet1!A1:C4", true);
        table.name = "myTable";
        return ctx.sync().then(function () {
            console.log("MyTable is Created!");

            //Create a new table binding for myTable
            Office.context.document.bindings.addFromNamedItemAsync("myTable", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    // If successful, add the event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
    });
    
    // When data in the table is changed, this event is triggered.
    function onBindingDataChanged(eventArgs) {
        Excel.run(function (ctx) {
            // Highlight the table in orange to indicate data changed.
            var fill = ctx.workbook.tables.getItem("myTable").getDataBodyRange().format.fill;
            fill.load("color");
            return ctx.sync().then(function () {
                if (fill.color != "Orange") {
                    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
 
                    console.log("The value in this table got changed!");
                }
                else
                    
            })
                .then(ctx.sync)
            .catch(function (error) {
                console.log(JSON.stringify(error));
            });
        });
    } 
})();
 


```



#### Beispiele
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Ruft ein Binding-Objekt anhand seiner Position im Elementarray ab.

#### Syntax
```js
bindingCollectionObject.getItemAt(index);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Index|number|Index-Wert des abzurufenden Objekts. Nullindiziert.|

#### Gibt Folgendes zurück:
[Bindung](binding.md)

#### Beispiele
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### load(param: object)
Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.

#### Syntax
```js
object.load(param);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder akzeptiert ein [loadOption](loadoption.md)-Objekt.|

#### Gibt 
void
### Eigenschaftszugriffsbeispiele

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
            console.log(bindings.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Abrufen der Anzahl von Bindungen

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
