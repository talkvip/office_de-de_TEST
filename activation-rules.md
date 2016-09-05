
# Aktivierungsregeln für Outlook-Add-Ins

Einige Typen von Add-Ins werden von Outlook aktiviert und in der Outlook-Benutzeroberfläche angezeigt, wenn die aktuelle Nachricht oder der aktuelle Termin die Aktivierungsregeln des Add-Ins erfüllt. Dies gilt für alle Ad-Ins, die das Manifestschema 1.1 verwenden, sowie für Add-Ins von benutzerdefinierten Bereichen. Anschließend kann der Benutzer das Add-In in der Outlook-Benutzeroberfläche auswählen, um es für das aktuelle Element zu starten.

Die folgende Abbildung zeigt Outlook-Add-Ins, die in der Add-In-Leiste für die Nachricht im Lesebereich aktiviert sind. 

![App-Leiste, die aktivierte Apps zum Lesen von E-Mails zeigt](../../../images/mod_off15_MailAppAppBar.png)


## Angeben von Aktivierungsregeln in einem Manifest


Wenn Sie die kontextbezogene Aktivierung verwenden möchten und in Outlook ein Add-In für bestimmte Bedingungen aktiviert werden soll, geben Sie Aktivierungsregeln im Add-In-Manifest an. Verwenden Sie hierfür eines dieser zwei **Rule**-Elemente:

- [Rule-Element (MailApp complexType)](../../../reference/manifest/rule.md) – Gibt eine einzelne Regel an.
- [Rule-Element (RuleCollection ComplexType)](#rule-element-rulecollection-complextype) – Kombiniert mehrere Regeln mithilfe von logischen Operationen.
    

 > **Hinweis:**  Das **Rule**-Element, das Sie zum Angeben einer einzelnen Regel verwenden, entspricht dem abstrakten komplexen Typ [Rule](../../../reference/manifest/rule.md). Jede der folgenden Regeltypen erweitert diesen abstrakten komplexen Typ **Rule**. Wenn Sie also eine einzelne Regel in einem Manifest angeben, müssen Sie das Attribut [xsi:type](http://www.w3.org/TR/xmlschema-1/) verwenden, um einen der folgenden Regeltypen genauer zu definieren. Die folgende Regel definiert z. B. eine [ItemIs](#itemis)-Regel: `<Rule xsi:type="ItemIs" ItemType="Message" />` Das **FormType**-Attribut bezieht sich auf Aktivierungsregeln in der Manifestdatei v1.1, ist jedoch in **VersionOverrides** v1.0 nicht definiert. Deshalb kann es nicht verwendet werden, wenn [ItemIs](#itemis) im Knoten **VersionOverrides** verwendet wird.

In der folgenden Tabelle werden die verfügbaren Regeltypen aufgeführt. Weitere Informationen finden Sie im Anschluss an die Tabelle und in den unter [Erstellen von Outlook-Add-Ins für Leseformulare](../../outlook/read-scenario.md) angegebenen Artikeln.


|**Regelname**|**Anwendbare Formulare**|**Beschreibung**|
|:-----|:-----|:-----|
|[ItemIs](#itemis)|Lesen, Verfassen, Benutzerdefinierter Bereich|Überprüft, ob das aktuelle Element den angegebenen Typ (Nachricht oder Termin) und (optional) die angegebene Elementnachrichtenklasse aufweist.|
|[ItemHasAttachment](#itemhasattachment)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob das ausgewählte Element über einen Anhang verfügt.|
|[ItemHasKnownEntity](#itemhasknownentity)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob das ausgewählte Element über eine oder mehrere allgemein bekannte Entitäten verfügt.Weitere Informationen: [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../../outlook/match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob die E-Mail-Adresse des Absenders, der Betreff und/oder der Textkörper des ausgewählten Elements mit einem regulären Ausdruck übereinstimmt.Weitere Informationen: [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](#rulecollection)|Lesen, Verfassen, Benutzerdefinierter Bereich|Kombiniert einen Satz von Regeln, damit Sie komplexere Regeln erstellen können.|

## ItemIs-Regel


Der komplexe Typ  **ItemIs** definiert eine Regel, deren Auswertung **true** ergibt, wenn das aktuelle Element dem Elementtyp und optional der Elementnachrichtenklasse entspricht, sofern diese in der Regel angegeben ist.

Geben Sie einen der folgenden Elementtypen in das Attribut  **ItemType** einer **ItemIs**-Regel ein. Sie können mehr als eine  **ItemIs**Regel in einem Manifest angeben. Der ItemType-SimpleType legt die Outlook-Elementtypen fest, die Outlook-Add-Ins unterstützen.



|**Wert**|**Beschreibung**|
|:-----|:-----|
|**Termin**|Gibt ein Element in einem Outlook-Kalender an. Dies kann eine Einladung sein, ein Besprechungselement, das beantwortet wurde und einen Organisator und Teilnehmer hat, oder ein Termin, der keinen Organisator oder Teilnehmer hat und einfach ein Element im Kalender darstellt.Entspricht der IPM.Appointment-Nachrichtenklassen in Outlook:|
|**Meldung**|Gibt eines der folgenden, typischerweise im Posteingang empfangenen Elemente an: <ul><li><p>Eine E-Mail-Nachricht. Entspricht der Nachrichtenklasse "IPM.Note" in Outlook.</p></li><li><p>Eine Besprechung, Antwort oder ein Abbruch. Entspricht den folgenden Nachrichtenklassen in Outlook:</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|
Das **FormType**-Attribut wird zum Angeben des Modus (Lesen oder Verfassen) verwendet, in dem das Add-In aktiviert werden sollte.


 > **Hinweis: **  Das ItemIs **FormType**-Attribut ist im Schema v1.1 und höher, aber nicht in **VersionOverrides** v1.0 definiert. Fügen Sie nicht das **FormType**-Attribut ein, wenn Sie Add-In-Befehle definieren.

Nach der Aktivierung eines Add-Ins können Sie mit der Eigenschaft [mailbox.item](../../../reference/outlook/Office.context.mailbox.item.md) das aktuell ausgewählte Element in Outlook und mit der Eigenschaft [item.itemType](../../../reference/outlook/Office.context.mailbox.item.md) die Art des aktuell ausgewählten Elements abrufen.

Optional können Sie mit dem  **ItemClass**-Attribut die Nachrichtenklasse des Elements angeben, und mit dem  **IncludeSubClasses**-Attribut können Sie angeben, ob die Regel " **true**" sein soll, wenn es sich bei dem Element um eine Unterklasse der angegebenen Klasse handelt.

Weitere Informationen zu Nachrichtenklassen finden Sie unter [Elementtypen und Meldungsklassen](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx).

Das folgende Beispiel ist eine  **ItemIs**-Regel, mit deren Hilfe Benutzer das Add-In in der Outlook-Add-In-Leiste sehen können, wenn der Benutzer eine Nachricht oder einen Termin liest.

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

Das folgende Beispiel ist eine  **ItemIs**-Regel, mit deren Hilfe Benutzer das Add-In in der Outlook-Add-In-Leiste sehen können, wenn der Benutzer eine Nachricht liest.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## ItemHasAttachment-Regel


Der komplexe Typ  **ItemHasAttachment** definiert eine Regel, die überprüft, ob das ausgewählte Element eine Anlage enthält.

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## ItemHasKnownEntity-Regel

Bevor ein Element für ein Add-In verfügbar gemacht wird, wird es vom Server untersucht, um festzustellen, ob der Betreff oder der Textkörper Text enthält, bei dem es sich um eine der bekannten Entitäten handeln könnte. Wird eine dieser Entitäten gefunden, wird sie in eine Sammlung bekannter Entitäten verschoben. Diese können Sie mit der  **getEntities**- oder der  **getEntitiesByType**-Methode dieses Elements aufrufen.

Mithilfe von **ItemHasKnownEntity** können Sie eine Regel angeben, mit der Ihr Add-In angezeigt wird, wenn eine Entität des angegebenen Typs in dem Element vorhanden ist. Im **EntityType**-Attribut einer **ItemHasKnownEntity**-Regel können die folgenden bekannten Entitäten angegeben werden:
-  Adresse
-  Kontakt
-  EmailAddress
-  MeetingSuggestion
-  PhoneNumber
-  TaskSuggestion
-  URL
    
Optional können Sie einen regulären Ausdruck in das  **RegularExpression**-Attribut integrieren, damit Ihr Add-In nur dann angezeigt wird, wenn eine Entität vorhanden ist, die mit dem regulären Ausdruck übereinstimmt. Um Übereinstimmungen mit regulären Ausrücken abzurufen, die in  **ItemHasKnownEntity**-Regeln angegeben sind, können Sie die Methode  **getRegExMatches** oder **getFilteredEntitiesByName** für das aktuell ausgewählte Outlook-Element verwenden.

Im folgenden Beispiel finden Sie eine Sammlung von  **Rule**-Elementen, mit denen das Add-in angezeigt wird, wenn sich eine der angegebenen bekannten Entitäten in der Nachricht befindet.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

Im folgenden Beispiel finden Sie eine  **ItemHasKnownEntity**-Regel mit einem  **RegularExpression**-Attribut, mit dem das Add-In aktiviert wird, wenn sich eine URL in der Nachricht befindet, die das Wort "contoso" enthält.


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Weitere Informationen zu Entitäten in Aktivierungsregeln finden Sie unter [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../../outlook/match-strings-in-an-item-as-well-known-entities.md).


## ItemHasRegularExpressionMatch-Regel


Mithilfe des komplexen Typs  **ItemHasRegularExpressionMatch** wird eine Regel definiert, in der ein regulärer Ausdruck verwendet wird, um die Inhalte der angegebenen Eigenschaft eines Elements abzugleichen. Wird in der angegebenen Eigenschaft des Elements Text gefunden, der mit dem regulären Ausdruck übereinstimmt, wird das Add-In von Outlook aktiviert und auf der Add-In-Leiste angezeigt. Mithilfe der Methode **getRegExMatches** oder **getRegExMatchesByName** des Objekts, das für das aktuell ausgewählte Element steht, können Sie Übereinstimmungen für den angegebenen regulären Ausdruck abrufen.

Im folgenden Beispiel ist ein  **ItemHasRegularExpressionMatch** dargestellt, der das Add-In aktiviert, wenn der Textkörper des ausgewählten Elements das Wort "apple", "banana" oder "coconut" enthält (Groß-/Kleinschreibung wird ignoriert).

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Weitere Informationen zur Verwendung der  **ItemHasRegularExpressionMatch**-Regel finden Sie unter [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).


## RuleCollection-Regel


Mithilfe des komplexen Typs " **RuleCollection**" werden mehrere Regeln in einer einzigen Regel kombiniert. Mit dem  **Mode**-Attribut können Sie angeben, ob die Regeln in der Sammlung mit einem logischen OR oder einem logischen AND kombiniert werden.

Wird ein logisches AND angegeben, muss ein Element mit allen in der Sammlung angegebenen Regeln übereinstimmen, damit das Add-In angezeigt wird. Wird ein logisches OR angegeben, wird das Add-in angezeigt, wenn ein Element mit einer der in der Sammlung angegebenen Regeln übereinstimmt.

Sie können  **RuleCollection**-Regeln kombinieren, um komplexe Regeln zu erstellen. Im folgenden Beispiel wird das Add-In aktiviert, wenn der Benutzer einen Termin oder eine Nachricht anzeigt und der Betreff oder der Textkörper des Elements eine Adresse enthält.

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

Mit dem folgenden Beispiel wird das Add-In aktiviert, wenn der Benutzer eine Nachricht verfasst oder wenn der Benutzer einen Termin anzeigt und der Betreff oder der Textkörper des Termins eine Adresse enthalten.

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## Einschränkungen für Regeln und reguläre Ausdrücke


Für eine zufriedenstellende Erfahrung mit Outlook-Add-Ins sollten Sie die Richtlinien für die Aktivierung und die API-Verwendung befolgen. In der folgenden Tabelle sind die allgemeinen Einschränkungen für reguläre Ausdrücke und Regeln dargestellt, es gibt jedoch spezifische Regeln für unterschiedliche Hosts. Weitere Informationen finden Sie unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](../../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md) und [Problembehandlung für die Aktivierung von Outlook-Add-Ins](../../outlook/troubleshoot-outlook-add-in-activation.md).

|**Add-In-Element**|**Richtlinien**|
|:-----|:-----|
|Manifestgröße|Nicht größer als 32 KB|
|Regeln|Nicht mehr als 15 Regeln.|
|ItemHasKnownEntity|Ein Outlook-Rich-Client wendet die Regel auf das erste MB des Elementtexts an, jedoch nicht auf den über diesen Grenzwert hinausgehenden Textkörper.|
|Reguläre Ausdrücke|Für die Regeln ItemHasKnownEntity oder ItemHasRegularExpressionMatch für alle Outlook-Hosts:<br><ul><li>Sie können nicht mehr als 5 reguläre Ausdrücke in Aktivierungsregeln in einem Outlook-Add-In angeben. Bei Überschreitung dieses Grenzwerts ist die Installation eines Outlook-Add-Ins nicht möglich.</li><li>Geben Sie reguläre Ausdrücke an, deren erwartete Ergebnisse vom Aufruf der <b>getRegExMatches</b>-Methode innerhalb der ersten 50 Übereinstimmungen zurückgegeben werden. </li><li>Geben Sie Look-Ahead-Assertionen in regulären Ausdrücken an, jedoch keine Look-Behind- (?<=text) und negativen Look-Behind-Assertionen (?<!text).</li><li>Geben Sie reguläre Ausdrücke an, deren Übereinstimmungen die Beschränkungen der nachstehenden Tabelle nicht überschreiten.<br/><br/><table><tr><th>Grenzwert für eine Übereinstimmung für einen regulären Ausdruck</th><th>Outlook-Rich-Clients</th><th>Outlook Web App für Geräte</th></tr><tr><td>Elementtext ist Nur-Text</td><td>1,5 KB</td><td>3 KB</td></tr><tr><td>Elementtext ist HTML</td><td>3 KB</td><td>3 KB</td></tr></table>|

## Weitere Ressourcen

- [Outlook-Add-Ins](../../outlook/outlook-add-ins.md)
- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](../../outlook/compose-scenario.md)
- [Einschränkungen für die Aktivierung und die JavaScript-API für Outlook-Add-Ins](../../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Elementtypen und Meldungsklassen](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
