# Action-Element
 Gibt die Aktion an, die durchgeführt werden soll, wenn der Benutzer ein [Schaltflächen](./button-control.md)- oder [Menü](./menu-control.md)steuerelement auswählt.
 
## Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [xsi:type](#xsi-type)  |  Ja  | Die Art der durchzuführenden Aktion|


## Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Gibt den Namen der auszuführenden Funktion an. |
|  [SourceLocation](#sourcelocation) |    Gibt den Speicherort der Quelldatei für diese Aktion an. |
  

## xsi:type
Dieses Attribut gibt die Art von Aktion an, die durchgeführt wird, wenn der Benutzer auf die Schaltfläche klickt. Das kann eine der folgenden Aktionen sein:
- ExecuteFunction
- ShowTaskpane

## FunctionName
Erforderliches Element, wenn **xsi:type** „ExecuteFunction“ ist. Gibt den Namen der auszuführenden Funktion an. Die Funktion ist in der Datei enthalten, die im Element [FunctionFile](./functionfile.md) angegeben ist.

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## SourceLocation
Erforderliches Element, wenn  **xsi:type** "ShowTaskpane" ist. Gibt den Speicherort der Quelldatei für diese Aktion an. Das Attribut **resid** muss auf den Wert des **id**-Attributs eines  **Url**-Elements im  [Urls](./resources.md#urls)-Element im  [Resources](./resources.md) -Element festgelegt werden.

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  
