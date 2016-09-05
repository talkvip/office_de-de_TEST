# Hosts-Element

Gibt die Office-Clientanwendung an, in der das Office-Add-In aktiviert wird. Enthält eine Sammlung von **Host**-Elementen mit den entsprechenden Einstellungen. 

Wenn im [VersionOverrides](./versionoverrides.md) Element des **Hosts**-Elements im übergeordneten Teil des Manifests. 

## Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Host](#host)    |  Ja   |  Beschreibt einen Host und die zugehörigen Einstellungen. |

> **Hinweis: ** Outlook erfordert, dass `Hosts` eine `Host`-Definition für `MailHost` enthalten.

---- 

## Host-Element
Gibt einen einzelnen Office-Anwendungstyp an, bei dem das Add-In aktiviert werden soll, z. B. „Dokument“, „Arbeitsmappe“, „Präsentation“, „Projekt“, „Postfach“ und „Notizbuch“.

### Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [xsi:type](#xsi-type)  |  Ja  | Beschreibt den Office-Host, für den diese Einstellungen gelten.|

### Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Ja   |  Definiert den betroffenen Formfaktor. |


### xsi:type
Steuert, auf welchen Office-Host (Word, Excel, PowerPoint, Outlook, OneNote) die enthaltenen Einstellungen angewendet werden. Bei dem Wert muss es sich um Folgendes handeln:

- `MailHost` (Outlook)    


## Beispiel für Hosts 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
