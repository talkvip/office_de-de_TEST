
# Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest
Verwenden Sie das  **VersionOverrides** -Element in Ihrem Outlook-Add-In-Manifest, um Add-In-Befehle zu definieren.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Zur Unterstützung von Add-In-Befehlen wurden einige zusätzliche Elemente zum Add-In-Manifest v1.1 innerhalb  **VersionOverrides** -Elements hinzugefügt. Wenn ein Manifest das **VersionOverrides**-Element enthält, verwenden Versionen von Outlook, die Add-In-Befehle unterstützen, die Informationen innerhalb dieses Elements, um das Add-In zu laden. Ältere Versionen von Outlook, die keine Add-In-Befehle unterstützen, ignorieren das Element und verwenden weiterhin die alten Elemente, wie in [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md) beschrieben.

Wenn die Clientanwendung den Knoten  **VersionOverrides** erkennt, wird der Add-In-Name im Menüband angezeigt und nicht im Bereich zum Lesen/Verfassen. Das Add-In wird nicht an beiden Orten angezeigt.


## VersionOverrides-Element

Das  **VersionOverrides** -Element ist der Stammknoten, der die Informationen für die Add-In-Befehle enthält. Er wird in Manifestschema v1. 1 oder höher unterstützt, ist jedoch im VersionOverrides-Schema v1. 0 definiert. Die Attribute für **VersionOverrides** sind im Folgenden dargestellt.


|
|
|**Attribut**|**Beschreibung**|
|:-----|:-----|
|**xmlns**| Erforderlich. Der Schemaspeicherort. Muss "http://schemas.microsoft.com/office/mailappversionoverrides" sein.|
|**xsi:type**|Erforderlich. Die Schemaversion. Die in diesem Thema beschriebene Version ist "VersionOverridesV1_0".|
In der folgenden Tabelle sind die untergeordneten Elemente von  **VersionOverrides** gezeigt.


|
|
|**Element**|**Beschreibung**|
|:-----|:-----|
|**Description**|Beschreibt das Add-In. Dadurch wird das  **Beschreibung** -Element in allen übergeordneten Teilen des Manifests überschrieben. Der Text der Beschreibung ist in einem untergeordneten Element des **LongString**-Element im  **Ressourcen** -Element enthalten. Das **resid**-Attribut des  **Description**-Elements wird auf en Wert des  **id**-Attributs des  **String**-Elements festgelegt, das den Text enthält.|
|**Requirements**|Gibt den Mindestanforderungssatz und die Version von Office.js an, den das Office-Add-In zum Aktivieren benötigt. Die Definition ist dieselbe wie in [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md). Dies überschreibt das Element  **Anforderungen** im übergeordneten Teil des Manifests.|
|**Hosts**|Erforderlich. Gibt eine Sammlung von Hosttypen und ihren Einstellungen an. Überschreibt das Element  **Hosts** im übergeordneten Teil des Manifests. Muss ein **xsi:type** -Attribut aufweisen, das auf "MailHost" festgelegt ist, und muss über den untergeordneten Knoten **FormFactor** verfügen.|
|**Resources**|Definiert eine Sammlung von Ressourcen (Zeichenfolgen, URLs und Bilder), auf die von anderen Elementen des Manifests verwiesen wird. Dies wird im Abschnitt [Ressources-Element](#ressources-element) weiter unten in diesem Thema beschrieben.|
Im Nachfolgenden sehen Sie ein Beispiel für  **VersionOverrides** mit untergeordneten Elementen.




```XML
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```


## FormFactor-Element

Das  **FormFactor** -Element gibt die Einstellungen für ein Add-In für einen bestimmten Formfaktor an. Es ist ein untergeordneter Knoten unter **Hosts** / **Host**. Derzeit kann das Element nur den Desktop ( **DesktopFormFactor**) angeben. Es enthält alle Add-In-Informationen für diesen Formfaktor mit Ausnahme des Knotens  **Ressourcen**.

Der Formfaktor enthält das Element  **FunctionFile** und ein oder mehrere **ExtensionPoint** -Elemente. Weitere Informationen finden Sie in den nachfolgenden Abschnitten [FunctionFile-Element](#functionfile-element) und [ExtensionPoint-Element](#extensionpoint-element). Im Folgenden ist ein Beispiel für  **FormFactor** mit untergeordneten Knoten zu sehen.




```XML
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="CustomPane">
            <!-- information on this extension point -->
          </ExtensionPoint> 

          <!-- possibly more ExtensionPoint elements -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```


## FunctionFile-Element


Das  **FunctionFile** -Element ist ein untergeordnetes Element unter **FormFactor**. Es gibt die Quellcodedatei für Vorgänge an, die ein Add-In über Add-In-Befehle verfügbar macht, die eine JavaScript-Funktion ausführen, anstatt eine Benutzeroberfläche anzuzeigen. Das **resid**-Attribut des  **FunctionFile**-Elements wird auf den Wert des  **id**-Attributs eines  **Url**-Elements im  **Resources**-Element festgelegt, das die URL zu einer HTML-Datei enthält, die alle JavaScript-Funktionen enhält oder lädt, die von Add-In-Befehlsschaltflächen ohne Benutzeroberfläche verwendet werden. Weitere Informationen finden Sie im Abschnitt [Steuerelemente für Schaltflächen](#steuerelemente-für-schaltflächen) dieses Artikels

Das JavaScript in der HTML-Datei, das vom  **FunctionFile**-Element angegeben wird, muss  `Office.initialize` aufrufen und benannte Funktionen definieren, die einen einzigen Parameter verwenden: `event`. Die Funktionen sollten die [item.notificationMessages](../reference/outlook/Office.context.mailbox.item.md%28Office.15%29.md)-API verwenden, um den Fortschritt, Erfolg oder Fehler für den Benutzer anzuzeigen. Sie sollten auch [event.completed (JavaScript API für Office)](http://msdn.microsoft.com/library/89db707a-a522-4a71-a830-801d0106527b%28Office.15%29.aspx) aufrufen, wenn der Vorgang abgeschlossen ist. Der Name der Funktionen wird im **FunctionName**-Element für Schaltflächen ohne Benutzeroberfläche verwendet.

Nachfolgend ein Beispiel für eine HTML-Datei, die eine trackMessage-Funktion definiert.




```
Office.intialize = function () {
    doAuth();
}
function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}

```


## ExtensionPoint-Element


Das  **ExtensionPoint** -Element definiert, wo ein Add-In eine Funktion bereitstellt. Es handelt sich um ein untergeordnetes Element unter **FormFactor**. Für jeden Formularfaktor können Entwickler **ExtensionPoint**-Elemente mit den folgenden  **xsi:type**-Werten festlegen:


- CustomPane 
    
- MessageReadCommandSurface 
    
- MessageComposeCommandSurface 
    
- AppointmentOrganizerCommandSurface 
    
- AppointmentAttendeeCommandSurface
    

### CustomPane

Der Erweiterungpunkt  **CustomPane** definiert ein Add-In, das aktiviert wird, wenn die angegebenen Regeln erfüllt werden. Er ist nur für das Leseformular vorgesehen und wird in einem horizontalen Bereich angezeigt. Im folgenden sind die Elemente von **CustomPane** gezeigt.


|
|
|**Element**|**Beschreibung**|
|:-----|:-----|
|**RequestedHeight**| Optional. Die erforderliche Höhe in Pixeln für den Anzeigebereich bei Ausführung auf einem Desktopcomputer. Sie kann zwischen 32 und 450 Pixeln liegen. Es ist dasselbe Element wie in Lese-Add-Ins (siehe[RequestedHeight-Element (ItemReadTabletMailAppSettings complexType) (App-Manifestschema v1.1)](http://msdn.microsoft.com/library/6296f5b0-3d5b-5ab9-eee9-55a7eb90f92c%28Office.15%29.aspx).|
|**SourceLocation**|Erforderlich. Die URL für die Quellcodedatei des Add-Ins. Dies bezieht sich auf ein  **Url**-Element im  **Resources** -Element.|
|**Rule**|Erforderlich. Die Regel oder Sammlung von Regeln, die angibt, wann das Add-In aktiviert wird. Das Element ist dasselbe wie das in [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md) definierte, mit der Ausnahme, dass die [ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)-Regel die folgenden Änderungen aufweist:  **ItemType** ist entweder "Message" oder "AppointmentAttendee", und es gibt kein **FormType** -Attribut. Weitere Informationen finden Sie unter [Benutzerdefinierter Bereich Outlook-Add-Ins](../outlook/custom-pane-outlook-add-ins.md) und [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md).|
|**DisableEntityHighlighting**|Optional. Gibt an, ob die Hervorhebung von Entitäten für dieses Mail-Add-In deaktiviert werden soll. |
 Im folgenden Beispiel wird ein benutzerdefinierter Bereich für Elemente definiert, die Nachrichten sind oder eine Anlage haben oder eine Adresse enthalten.




```
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

Auf dem Menüband geben Entwickler die Registerkarte und Gruppe für ihre Add-In-Befehle an. Dies kann entweder die Standardregisterkarte ( **Start**,  **Nachricht** oder **Besprechung**) oder eine vom Add-In definierte benutzerdefinierte Registerkarte sein. Beim Hinzufügen zur Standardregisterkarte ist dies auf eine Gruppe pro Add-In beschränkt. Bei benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente beschränkt, unabhängig davon, auf welcher Registerkarte sie angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte beschränkt.

Ein Beispiel für eine Gruppe auf der Registerkarte des Menübands ist wie folgt.




```XML
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
      <Label resid="residTemplateManagement" />
      <Control xsi:type="Button" id="Button1" >
       <!-- information on the control -->
      </Control>
       <!-- other controls, as needed -->
    </Group>
  </OfficeTab>
</ExtensionPoint>

```

Wobei Folgendes gilt:


|
|
|**Element**|**Beschreibung**|
|:-----|:-----|
|**OfficeTab**|Erforderlich. Die vorab vorhandene Registerkarte, die verwendet werden soll. Derzeit kann das  **id**-Attribut nur „TabDefault" sein.|
|**Group**|Eine Gruppe von Erweiterungspunkten der Benutzeroberfläche auf einer Registerkarte. Eine Gruppe kann über bis zu sechs Steuerelemente verfügen.Das Attribut  **Id** ist erforderlich. Es ist eine Zeichenfolge mit maximal 125 Zeichen.|
|**Label**|Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**Control**|Eine Gruppe erfordert mindestens ein Steuerelement. Derzeit werden nur Schaltflächen und Menüs unterstützt. Weitere Informatinen finden Sie in den nachfolgenden Abschnitten [Steuerelemente für Schaltflächen](#steuerelemente-für-schaltflächen) und[Steuerelemente für Menüs (Dropdownschaltfläche)](#steuerelemente-für-menüs-dropdownschaltfläche).|
Entwickler können außerdem eine benutzerdefinierte Registerkarte im Menüband erstellen, indem sie das Element  **CustomTab** verwenden, wie im folgenden Beispiel gezeigt.




```XML
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
      <Label resid="residCustomTabGroupLabel"/>
      <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

Wobei Folgendes gilt:


|
|
|**Element**|**Beschreibung**|
|:-----|:-----|
|**CustomTab**|Erforderlich. Das Attribut  **id** muss im Manifest eindeutig sein.|
|**Group**|Eine Gruppe von Erweiterungspunkten der Benutzeroberfläche auf einer Registerkarte. Eine Gruppe kann über bis zu sechs Steuerelemente verfügen.Das Attribut  **id** ist erforderlich. Es handelt sich um eine Zeichenfolge mit maximal 125 Zeichen.|
|**Label**|Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**Control**|Eine Gruppe erfordert mindestens ein Steuerelement. Derzeit werden nur Schaltflächen und Menüs unterstützt. Weitere Informatinoen finden Sie in den nachfolgenden Abschnitten [Steuerelemente für Schaltflächen](#steuerelemente-für-schaltflächen) und[Steuerelemente für Menüs (Dropdownschaltfläche)](#steuerelemente-für-menüs-dropdownschaltfläche).|
|**Label**|Erforderlich. Die Beschriftung der benutzerdefinierten Registerkarte. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|

#### Steuerelemente für Schaltflächen


Eine Schaltfläche führt eine einzelne Aktion durch, wenn der Benutzer auf die Schaltfläche klickt. Sie kann entweder eine Funktion ausführen oder einen Aufgabengereich anzeigen.

Das Steuerelement für Schaltflächen sieht folgendermaßen aus:




```
<Control xsi:type="Button" id="<choose a descriptive name>" >
  <!-- include button elements, as described in the following table -->
</Control>
```

Dabei ist das Attribut  **Id** eine Zeichenfolge mit maximal 125 Zeichen. Die Schaltflächenelemente sind in der folgenden Tabelle beschrieben.


|
|
|**Schaltflächenelemente**|**Beschreibung**|
|:-----|:-----|
|**Label**|Erforderlich. Der Text für die Schaltfäche. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**Supertip**|Erforderlich. Der SuperTip für diese Schaltfläche, der mit Folgendem definiert wird
|||
|:-----|:-----|
|**Title**|Erforderlich. Der Text für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**Description**|Erforderlich. Die Beschreibung für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **LongStrings**-Element im  **Resources** -Element festgelegt werden.|
|
|**Icon**|Erforderlich. Enthält die  **Image**-Elemente für die Schaltfläche. 
|||
|:-----|:-----|
|**Image**|Ein Bild für die Schaltfläche. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines  **Image**-Elements im  **Images**-Element im  **Resources** -Element festgelegt werden. Das Attribut **size** gibt die Größe des Bilds in Pixel an. Drei Bildgrößen sind erforderlich (16, 32 und 80 Pixel), es werden aber fünf weitere Größen unterstützt (20, 24, 40, 48 und 64 Pixel).|
|
|**Action**|Erforderlich. Gibt die Aktion an, die durchgeführt werden soll, wenn der Benutzer auf die Schaltfläche klickt. Das Element wird durch Folgendes definiert.
|||
|:-----|:-----|
|**xsi:type**|Dieses Attribut gibt die Art von Aktion an, die durchgeführt wird, wenn der Benutzer auf die Schaltfläche klickt. Das kann eine der folgenden Aktionen sein:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>"ExecuteFunction"</p></li><li><p>"ShowTaskpane"</p></li></ul>|
|**FunctionName**|Erforderliches Element, wenn  **xsi:type** "ExecuteFunction" ist. Gibt den Namen der auszuführenden Funktion an. Die Funktion ist in der Datei enthalten, die im Element **FunctionFile** angegeben ist.|
|**SourceLocation**|Erforderliches Element, wenn  **xsi:type** "ShowTaskpane" ist. Gibt den Speicherort der Quelldatei für diese Aktion an. Das Attribut **resid** muss auf den Wert des **id**-Attributs eines  **Url**-Elements im  **Urls**-Element im  **Resources** -Element festgelegt werden.|
|
Nachfolgend sehen Sie ein Beispiel für eine  _Schaltfläche ohne Benutzeroberfläche_, die eine Funktion namens `getSubject` ausführt.




```
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

Ein Steuerelement für eine  _Aufgabenbereichsschaltfläche_ ist eine Schaltfläche, die einen Aufgabenbereich startet. Aufgabenbereichsschaltflächen unterstützen keine Umschaltung. Es folgt ein Beispiel für eine Aufgabenbereichsschaltfläche.




```
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


#### Steuerelemente für Menüs (Dropdownschaltfläche)


Ein Menü definiert eine statische Liste mit Optionen. Jedes Menüelement führt eine Funktion aus oder zeigt einen Aufgabenbereich an. Untermenüs werden nicht unterstützt. 

Die Syntax für das Menüsteuerelement lautet wie folgt:




```
<Control xsi:type="Menu" id="<choose a descriptive name>" >
  <!-- include menu elements, as described in the following table -->
</Control>
```

Dabei ist das Attribut  **Id** eine Zeichenfolge mit maximal 125 Zeichen. Die Menüelemente sind in der folgenden Tabelle beschrieben.


|
|
|**Menüelemente**|**Beschreibung**|
|:-----|:-----|
|**Label**|Erforderlich. Der Text für das Menü. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**SuperTip**|Erforderlich. Der SuperTip für das Menü, der mit Folgendem definiert wird
|||
|:-----|:-----|
|**Title**|Erforderlich. Der Text für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **ShortStrings**-Element im  **Resources** -Element festgelegt werden.|
|**Description**|Erforderlich. Die Beschreibung für den SuperTip. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines Elements im  **LongStrings**-Element im  **Resources** -Element festgelegt werden.|
|
|**Icon**|Erforderlich. Enthält die  **Image**-Elemente für das Menü. 
|||
|:-----|:-----|
|**Image**|Ein Bild für das Menü. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines  **Image**-Elements im  **Images**-Element im  **Resources** -Element festgelegt werden. Das Attribut **size** gibt die Größe des Bilds in Pixel an. Drei Bildgrößen sind erforderlich (16, 32 und 80 Pixel), es werden aber fünf weitere Größen unterstützt (20, 24, 40, 48 und 64 Pixel).|
|
|**Items**|Erforderlich. Enthält die  **Item**-Elemente für das Menü. Jedes  **Item**-Element enthält dieselben untergeordneten Elemente wie ein [Steuerelemente für Schaltflächen](#steuerelemente-für-schaltflächen).|


Nachfolgend finden Sie ein Beispiel für ein Menü.




```
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


### MessageComposeCommandSurface

Platziert Schaltflächen für Add-Ins, die Mailerstellformulare verwenden, im Menüband. Es wird auf dieselbe Weise wie MessageReadCommandSurface definiert.


### AppointmentOrganizerCommandSurface

Platziert Schaltflächen für das Formular, das dem Organisator der Besprechung angezeigt wird, im Menüband. Es wird auf dieselbe Weise wie MessageReadCommandSurface definiert.


### AppointmentAttendeeCommandSurface

Platziert Schaltflächen für das Formular, das dem Teilnehmer der Besprechung angezeigt wird, im Menüband. Es wird auf dieselbe Weise wie MessageReadCommandSurface definiert.


## Ressources-Element


Das Element  **Ressourcen** enthält Symbole, Zeichenfolgen und URLs für den Knoten **VersionOverrides**. Ein Manifestelement gibt eine Ressource mithilfe der **Id** der Ressource an. Dadurch bleibt die Größe des Manifests verwaltbar, insbesondere wenn Ressourcen über Versionen für unterschiedliche Gebietsschemas verfügen. Eine **Id** hat maximal 32 Zeichen.

Der Knoten  **Ressourcen** definiert die folgenden Ressourcen. Jede Ressource kann über ein oder mehrere untergeordnete **Override** -Elemente verfügen, um eine Ressource für bestimmte Gebietsschemas zu definieren.


|
|
|**Ressource**|**Beschreibung**|
|:-----|:-----|
|**Images**/ **Image**|Gibt die HTTPS-URL zu einem Bild für ein Symbol an. Jedes Symbol muss drei  **Image**-Elemente aufweisen, jeweils eines für die drei obligatorischen Größen: 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>16 x 16</p></li><li><p>32 x 32</p></li><li><p>80 x 80</p></li></ul>Die folgenden zusätzlichen Größen werden ebenfalls unterstützt, sind jedoch nicht erforderlich:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>20 x 20</p></li><li><p>24 x 24</p></li><li><p>40 x 40</p></li><li><p>48 x 48</p></li><li><p>64 x 64</p></li></ul>|
|**Urls**/ **Url**|Stellt eine HTTPS-URL bereit. Eine URL kann maximal 2.048 Zeichen lang sein. |
|**ShortStrings**/ **String**|Der Text für die Elemente  **Bezeichnung** und **Titel**. Jede **String** enthält maximal 125 Zeichen|
|**LongStrings**/ **String**|Der Text für die Attribute  **Description**. Jedes **String** enthält maximal 250 Zeichen.|

(../../images  Beim Definieren von Ressourcen sollten Sie die folgenden Anforderungen beachten:

Im folgenden finden Sie ein Beispiel für das Element  **Ressourcen**.




```
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="??? ??????? ????????" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="???? ??? ????? ??????? ?? ???????." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>

```


## Regeländerungen


Die folgenden Änderungen betreffen die Regeln im Manifest.


- Aktivierungsregeln befinden sich jetzt in jedem Einstiegspunkt.
    
- [ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) wird geändert, damit **ItemType** „Message" oder „AppointmentAttendee" entspricht und kein **FormType** -Attribut vorhanden ist.
    
- [ItemHasKnownEntity](http://msdn.microsoft.com/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) wird geändert, um eine Zeichenfolge für den Entitätstyp statt eine Enumeration zu akzeptieren.
    

## Beispielmanifest


Ein vollständiges Beispielmanifest finden Sie unter [command-demo](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml%28Office.15%29.aspx) auf GitHub.


## Zusätzliche Ressourcen



- [Add-In-Befehle für Outlook](../outlook/add-in-commands-for-outlook.md)
    
- [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md)
    
- [command-demo sample](https://github.com/jasonjoh/command-demo.aspx)
    
