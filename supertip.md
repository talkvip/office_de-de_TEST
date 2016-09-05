## Multiinfo
Definiert eine umfangreiche QuickInfo (Titel und Beschreibung). Es wird von [Button](./button.md)- und [Menu](./menu-control.md)Steuerelementen verwendet. 

## Untergeordnete Elemente
|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Position](#position)        | Ja |   Der Text für die Multiinfo.         |
|  [Beschreibung](#beschreibung)  | Ja |  Die Beschreibung für die Multiinfo.    |

## Title
Erforderlich. Der Text für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  [Resources](./resources.md#shortstrings) -Element festgelegt werden.

## Beschreibung
Erforderlich. Die Beschreibung für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **LongStrings**-Element im  [Resources](./resources.md#longstrings) -Element festgelegt werden.

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```