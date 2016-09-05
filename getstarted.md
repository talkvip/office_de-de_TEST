# GetStarted-Element

Stellt Informationen bereit, die von der Beschriftung verwendet werden, die angezeigt wird, wenn das Add-In auf Word-, Excel-, PowerPoint- und OneNote-Hosts installiert wird. Das **GetStarted**-Element ist ein untergeordnetes Element von [FormFactor](./formfactor.md).

## Untergeordnete Elemente

| Element                       | Erforderlich | Beschreibung                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Position](#position)               | Ja      | Definiert, wo ein Add-In Funktionen verfügbar macht.     |
| [Beschreibung](#beschreibung)   | Ja      | Eine URL zu einer Datei, die JavaScript-Funktionen enthält.|
| [LearnMoreUrl](#learnmoreurl) | Nein       | Eine URL zu einer Seite, auf der das Add-In ausführlich erläutert wird.   |


## Title 
Erforderlich. Der Titel, der für den oberen Rand der Beschriftung verwendet wird. Das**resid**-Attribut verweist auf eine gültige ID im [ShortStrings](./resources.md#shortstrings)-Element im Abschnitt [Ressourcen](./resources.md).

## Beschreibung
Erforderlich. Die Beschreibung/der Textinhalt der Beschriftung. Das**resid**-Attribut verweist auf eine gültige ID im [LongStrings](./resources.md#longstrings)-Element im Abschnitt [Ressourcen](./resources.md).

## LearnMoreUrl
Erforderlich. Die URL zu einer Seite, auf der der Benutzer mehr über das Add-In erfahren kann. Das**resid**-Attribut verweist auf eine gültige ID im [Urls](./resources.md#urls)-Element im Abschnitt [Ressourcen](./resources.md).

> **HINWEIS:** **LearnMoreUrl** wird derzeit auf Word-, Excel- oder PowerPoint-Clients nicht wiedergegeben. Es wird empfohlen, dass Sie diese URL für alle Clients hinzufügen, damit die URL dargestellt wird, sobald sie verfügbar ist. 
