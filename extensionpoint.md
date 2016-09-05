# ExtensionPoint-Element

 Definiert, wo ein Add-In Funktionen in der Office-Benutzeroberfläche verfügbar macht. Das **ExtensionPoint**-Element ist ein untergeordnetes Element von [FormFactor](./formfactor.md). 

## Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Ja  | Der Typ des Erweiterungspunkts, der definiert wird.|


## Erweitertungspunkte für Word-, Excel-, PowerPoint und OneNote-Add-In-Befehle

- **PrimaryCommandSurface** Das Menüband in Office.
- **ContextMenu** Das Kontextmenü, das angezeigt wird, wenn Sie mit der rechten Maustaste auf die Office-Benutzeroberfläche klicken.

In den folgenden Beispielen wird gezeigt, wie Sie das  **ExtensionPoint**-Element mit  **PrimaryCommandSurface**- und  **ContextMenu**-Attributwerten verwenden, und Sie sehen die untergeordneten Elemente, die für jedes Element verwendet werden sollten.


 >**Wichtig**  Geben Sie bei Elementen, die ein ID-Attribut enthalten, eine eindeutige ID an. Es wird empfohlen, dass Sie den Namen Ihres Unternehmens zusammen mit Ihrer ID verwenden. Verwenden Sie zum Beispiel folgendes Format:<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**Untergeordnete Elemente**
 
|**Element**|**Beschreibung**|
|:-----|:-----|
|**CustomTab**|Erforderlich, wenn Sie eine benutzerdefinierte Registerkarte (unter Verwendung von  **PrimaryCommandSurface**) zum Menüband hinzufügen möchten. Wenn Sie das  **CustomTab**-Element verwenden, können Sie das  **OfficeTab**-Element nicht verwenden. Das  **ID** -Attribut ist erforderlich.|
|**OfficeTab**|Erforderlich, wenn Sie eine standardmäßige Office-Menübandregisterkarte (unter Verwendung von **PrimaryCommandSurface**) erweitern möchten. Wenn Sie das **OfficeTab**-Element verwenden, können Sie das **CustomTab**-Element nicht verwenden. Weitere Informationen finden Sie unter [OfficeTab](officetab.md)|
|**OfficeMenu**|Erforderlich, wenn Sie Add-In-Befehle zu einem Standardkontextmenü (unter Verwendung von **ContextMenu**) hinzufügen. Das **id**-Attribut muss festgelegt werden auf: <br/> - **ContextMenuText** für Excel oder Word. Zeigt das Element im Kontextmenü an, wenn Text ausgewählt wird und der Benutzer dann mit der rechten Maustaste auf den ausgewählten Text klickt. <br/> - **ContextMenuCell** für Excel. Zeigt das Element im Kontextmenü an, wenn der Benutzer mit der rechten Maustaste in eine Zelle in dem Arbeitsblatt klickt.|
|**Gruppe**|Eine Gruppe von Erweiterungspunkten der Benutzeroberfläche auf einer Registerkarte.Eine Gruppe kann über bis zu sechs Steuerelemente verfügen. Das **id** -Attribut ist erforderlich. Es handelt sich um eine Zeichenfolge mit maximal 125 Zeichen.|
|**Label**|Erforderlich. Die Bezeichnung der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festegelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **ShortStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.|
|**Symbol**|Erforderlich. Gibt das Gruppensymbol an, das auf kleinen Geräten verwendet wird, oder wenn zu viele Schaltflächen angezeigt werden. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **Image**-Elements festgelegt werden. Das  **Image**-Element ist ein untergeordnetes Element vom  **Images**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist. Das **size**-Attribut gibt die Größe des Bilds in Pixeln an. Es sind drei Bildgrößen erforderlich: 16, 32 und 80. Es werden auch fünf optionale Größen unterstützt: 20, 24, 40, 48 und 64.|
|**QuickInfo**|Optional. Die QuickInfo der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festgelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **LongStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.|
|**Control**|Für jede Gruppe ist mindestens ein Steuerelement erforderlich. Ein  **Control**-Element kann entweder eine  **Button** oder ein **Menu** sein. Verwenden Sie **Menu**, um eine Dropdownliste von Steuerelemente für Schaltflächen anzugeben. Derzeit werden nur Schaltflächen und Menüs unterstützt.Weitere Informationen finden Sie in den Abschnitten [Steuerelemente für Schaltflächen](#button-controls) und [Menüsteuerelemente](#menu-controls).<br/>**Hinweis** Zur Erleichterung der Fehlerbehebung wird empfohlen, dass das **Control**-Element und die zugehörigen untergeordneten **Resources**-Elemente jeweils einzeln hinzugefügt werden.

## Erweiterungspunkte für Outlook-Add-In-Befehle

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Kann nur im [DesktopFormFactor](./formfactor.md) verwendet werden.)

### CustomPane

Der CustomPane-Erweiterungspunkt definiert ein Add-In, das aktiviert wird, wenn die angegebenen Regeln erfüllt werden. Er ist nur für das Leseformular vorgesehen und wird in einem horizontalen Bereich angezeigt. 

**Untergeordnete Elemente**

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **RequestedHeight** | Nein |  Die erforderliche Höhe in Pixeln für den Anzeigebereich bei Ausführung auf einem Desktopcomputer. Sie kann zwischen 32 auf 450 Pixeln liegen.  |
|  **SourceLocation**  | Ja |  Die URL für die Quellcodedatei des Add-Ins. Dies bezieht sich auf ein **Url**-Element im [Ressourcen](./resources.md)-Element.  |
|  **Regel**  | Ja |  Die Regel oder Sammlung von Regeln, die angibt, wann das Add-In aktiviert wird. Weitere Informationen finden Sie unter [Aktivierungsregeln für Outlook-Add-Ins](../../outlook/manifests/activation-rules.md). |
|  **DisableEntityHighlighting**  | Nein |  Gibt an, ob die Hervorhebung von Entitäten deaktiviert werden soll. |


#### CustomPane-Beispiel
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
Mit diesem Erweiterungspunkt werden Schaltflächen für die Ansicht gelesener Mails auf der Befehlsoberfläche platziert. In Outlook Desktop wird das Element im Menüband angezeigt.

**Untergeordnete Elemente**

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](./customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
Dieser Erweiterungspunkt platziert Schaltflächen für Add-Ins, die Mailformulare zum Verfassen verwenden, im Menüband. 

**Untergeordnete Elemente**

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](./customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab-Beispiel

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Organisator der Besprechung angezeigt wird, im Menüband. 

**Untergeordnete Elemente**

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](./customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Teilnehmer der Besprechung angezeigt wird, im Menüband. 

**Untergeordnete Elemente**

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](./customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

Dieser Erweiterungspunkt platziert Schaltflächen für die Modulerweiterung im Menüband. 

**Untergeordnete Elemente**

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](./customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

