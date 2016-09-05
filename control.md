# Control-Element

Definiert eine JavaScript-Funktion, die eine Aktion ausführt oder einen Aufgabenbereich startet. Ein  **Control**-Element kann entweder eine Schaltfläche oder eine Menüoption sein. Mindestens ein **Control**-Element muss im [Group](group.md)-Element enthalten sein.

## Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|**xsi:type**|Ja|Der Typ des Steuerelements, das definiert wird. Kann eine Schaltfläche oder ein Menüs sein.|
|**id**|Nein|Die ID des Steuerelements. Darf maximal 125 Zeichen lang sein.|

## Schaltflächensteuerelement

Eine Schaltfläche führt bei Auswahl durch den Benutzer eine einzelne Aktion aus. Sie kann eine Funktion ausführen oder einen Aufgabenbereich anzeigen. Jedes Schaltflächen-Steuerelement muss eine für das Manifest eindeutige `id` aufweisen. 

### Untergeordnete Elemente
|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **Bezeichnung**     | Ja |  Der Text für die Schaltfläche. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  [ShortStrings](./resources.md#shortstrings)-Element im  [Resources](./resources.md)-Element festgelegt werden.        |
|  **Tooltip**  |Nein|Die QuickInfo für die Schaltfläche. Das Attribut **resid**  muss auf den Wert des Attributs **id** eines **String**-Elements festgelegt sein. Das Element **String** ist ein untergeordnetes Element des Elements **LongStrings**, und dieses ist ein untergeordnetes Element des Elements [Resources](resource.md).|     
|  [Multiinfo](./supertip.md)  | Ja |  Die Multiinfo für die Schaltfläche.    |
|  [Symbol](./icon.md)      | Ja |  Ein Bild für die Schaltfläche.         |
|  [Aktion](./action.md)    | Ja |  Gibt die auszuführende Aktion an.  |



```XML
        <!-- Define a control that calls a JavaScript function. -->

                 <Control xsi:type="Button" id="Button1Id1">
                  <Label resid="residLabel" />
                  <Tooltip resid="residToolTip" />
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getData</FunctionName>
                  </Action>
                </Control>


                <!-- Define a control that shows a task pane. -->

                <Control xsi:type="Button" id="Button2Id1">
                  <Label resid="residLabel2" />
                  <Tooltip resid="residToolTip" />
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon2_32x32" />
                    <bt:Image size="32" resid="icon2_32x32" />
                    <bt:Image size="80" resid="icon2_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Control>
```

### Beispiel für ExecuteFunction-Schaltfläche

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### Beispiel für ShowTaskpane-Schaltfläche

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
## Steuerelemente für Menüs (Dropdownschaltfläche)

Ein Menü definiert eine statische Liste mit Optionen. Jedes Menüelement führt eine Funktion aus oder zeigt einen Aufgabenbereich an. Untermenüs werden nicht unterstützt. 

Bei Verwendung mit einem **PrimaryCommandSurface**- oder **ContextMenu**[-Erweiterungspunkt](extensionpoint.md) definiert das Steuerelement Folgendes:

- Ein Menüelement auf der Stammebene.

- Eine Liste der Untermenüelemente.

Bei Verwendung mit  **PrimaryCommandSurface** zeigt das Menüelement auf der Stammebene eine Schaltfläche auf dem Menüband an. Wenn die Schaltfläche ausgewählt wird, wird das Untermenü als eine Dropdownliste angezeigt. Bei Verwendung mit **ContextMenu** wird ein Menüelement mit einem Untermenü im Kontextmenü eingefügt. In beiden Fällen können einzelne Untermenüelemente entweder eine JavaScript-Funktion ausführen oder einen Aufgabenbereich anzeigen. Derzeit wird nur eine Ebene von Untermenüs unterstützt.

Das folgende Beispiel zeigt, wie ein Menüelement mit zwei Untermenüelementen definiert wird. Das erste Untermenüelement zeigt einen Aufgabenbereich an, das zweite Untermenüelement führt eine JavaScript-Funktion aus.

```xml
<Control xsi:type="Menu" id="TestMenu2">
              <Label resid="residLabel3" />
              <Tooltip resid="residToolTip" />
              <Supertip>
                <Title resid="residLabel" />
                <Description resid="residToolTip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Items>
                <Item id="showGallery2">
                  <Label resid="residLabel3"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Item>
              <Item id="showGallery3">
                  <Label resid="residLabel5"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon4_32x32" />
                    <bt:Image size="32" resid="icon4_32x32" />
                    <bt:Image size="80" resid="icon4_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getButton</FunctionName>
                  </Action>
                </Item>
              </Items>
            </Control>

```

### Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **Bezeichnung**     | Ja |  Der Text für die Schaltfläche. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  [ShortStrings](./resources.md#shortstrings)-Element im  [Resources](./resources.md)-Element festgelegt werden.      |
|  **Tooltip**  |Nein|Die QuickInfo für die Schaltfläche. Das Attribut **resid**  muss auf den Wert des Attributs **id** eines **String**-Elements festgelegt sein. Das Element **String** ist ein untergeordnetes Element des Elements **LongStrings**, und dieses ist ein untergeordnetes Element des Elements [Resources](resource.md).|     
|  [Multiinfo](./supertip.md)  | Ja |  Die Multiinfo für diese Schaltfläche.    |
|  [Symbol](./icon.md)      | Ja |  Ein Bild für die Schaltfläche.         |
|  [Elemente](#elemente)     | Ja |  Eine Auflistung von Schaltflächen, die innerhalb des Menüs angezeigt werden sollen. Enthält die  **Item**-Elemente für jedes Untermenüelement. Jedes **Item**-Element enthält die untergeordneten Elemente eines [Schaltflächensteuerelements](#button-control).|


### Beispiele für Menüsteuerelemente

```xml
<Control xsi:type="Menu" id="TestMenu2">
              <Label resid="residLabel3" />
              <Tooltip resid="residToolTip" />
              <Supertip>
                <Title resid="residLabel" />
                <Description resid="residToolTip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Items>
                <Item id="showGallery2">
                  <Label resid="residLabel3"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Item>
              <Item id="showGallery3">
                  <Label resid="residLabel5"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon4_32x32" />
                    <bt:Image size="32" resid="icon4_32x32" />
                    <bt:Image size="80" resid="icon4_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getButton</FunctionName>
                  </Action>
                </Item>
              </Items>
            </Control>

```


```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
