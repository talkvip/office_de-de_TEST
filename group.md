# Group-Element
Definiert eine Gruppe von UI-Erweiterungspunkten auf einer Registerkarte.  In benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente begrenzt, unabhängig davon, in welcher Registerkarte es angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte begrenzt.

## Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [id](#id)  |  Ja  | Eine eindeutige ID für die Gruppe.|

## Untergeordnete Elemente
|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Bezeichnung](#bezeichnung)      | Ja |  Die Bezeichnung für CustomTab oder eine Gruppe  |
|  [Control](#control)    | Ja |  Auflistung von einem oder mehreren Control-Objekten.  |

## id-Attribut
Erforderlich. Eindeutiger Bezeichner für die Gruppe. Es ist eine Zeichenfolge mit maximal 125 Zeichen. Dies muss im Manifest eindeutig sein, andernfalls wird die Gruppe nicht gerendert.

## Label 
Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  [ShortStrings](./resources.md#shortstrings)-Element im  [Resources](./resources.md) -Element festgelegt werden.

## Control
Eine Gruppe erfordert mindestens ein Steuerelement. Derzeit werden nur [Schaltflächen](./control.md#button-control) und [Menüs](./menu.md#menu-control)unterstützt. 

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```