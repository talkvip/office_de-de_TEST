# InlinePictureCollection-Objekt (JavaScript-API für Word)

Enthält eine Sammlung von [InlinePicture](inlinepicture.md)-Objekten.

__Gilt für: Word 2016, Word für iPad, Word für Mac_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|Items|[InlinePicture[]](inlinepicture.md)|Eine Sammlung von InlinePicture-Objekten. Schreibgeschützt.|

## Beziehungen
Keine


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

## Supportdetails
Verwenden Sie den [Anforderungssatz](../office-add-in-requirement-sets.md) in Laufzeitüberprüfungen, um sicherzustellen, dass die Anwendung von der Hostversion von Word unterstützt wird. Weitere Informationen zur Office-Hostanwendung und den Serveranforderungen finden Sie unter [Anforderungen für die Ausführung von Office-Add-Ins](../../docs/overview/requirements-for-running-office-add-ins.md).
