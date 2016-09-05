# OfficeExtension.Error-Objekt (JavaScript-API für Excel)

Stellt Fehler dar, die bei Verwendung der JavaScript-API für Excel auftreten.

_Gilt für: Excel 2016, Excel Online, Excel für iOS, Office 2016_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|code|string|Ruft einen Wert ab, der den Typ des Fehlers angibt. Der Wert kann "AccessDenied", "ActivityLimitReached", "BadPassword", "GeneralException", "InsertDeleteConflict", "InvalidArgument", "InvalidBinding", "InvalidOperation", "InvalidReference", "InvalidSelection", "ItemAlreadyExists", "ItemNotFound", "NotImplemented" oder "UnsupportedOperation" sein. |
|debugInfo|string|Ruft einen Wert ab, der angibt, was bei Auftreten des Fehlers geschah. Dieser Wert ist nur für die Verwendung bei der Entwicklung/beim Debuggen vorgesehen.  |
|message |string| Ruft eine lokalisierte lesbare Zeichenfolge ab, die dem Fehlercode entspricht.|
|name |string| Ruft einen Wert ab, der immer "OfficeExtension.Error" ist. |
|traceMessages |string[]| Ruft ein Array von Werten ab, die den mit context.trace() festgelegten Instrumentierungsmeldungen entsprechen. |

## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Gibt die Werte für Fehlercode und Meldung im folgenden Format zurück: "{0}: {1}", code, message.|

## Methodendetails

### toString()
Gibt die Werte für Fehlercode und Meldung im folgenden Format zurück: "{0}: {1}", code, message.

#### Syntax
```js
error.toString()
```

#### Parameter
None.

#### Gibt 
string
