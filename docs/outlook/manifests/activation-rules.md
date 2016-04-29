
# Aktivierungsregeln für Outlook-Add-Ins
Definieren Sie Aktivierungsregeln in einer Outlook-Add-In-Manifestdatei, um den Kontext festzulegen, in dem Outlook das Add-In auf der Benutzeroberfläche anzeigen würde.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Aktivierung durch angegebene Regeln


Einige Typen von Add-Ins werden von Outlook aktiviert und in der Outlook-Benutzeroberfläche angezeigt, wenn die aktuelle Nachricht oder der aktuelle Termin die Aktivierungsregeln des Add-Ins erfüllt. Dies gilt für alle Ad-Ins, die das Manifestschema 1.1 verwenden, sowie für Add-Ins von benutzerdefinierten Bereichen. Anschließend kann der Benutzer das Add-In in der Outlook-Benutzeroberfläche auswählen, um es für das aktuelle Element zu starten.

Die folgende Abbildung zeigt Outlook-Add-Ins, die in der Add-In-Leiste für die Nachricht im Lesebereich aktiviert sind. 


**Die Add-In-Leiste zeigt zwei aktivierte Outlook-Add-Ins für die Nachricht im Leseformular des Lesebereichs.**

![App-Leiste, die aktivierte Apps zum Lesen von E-Mails zeigt](../../images/mod_off15_MailAppAppBar.png)


## Angeben von Aktivierungsregeln in einem Manifest


Wenn Sie die kontextbezogene Aktivierung verwenden möchten und in Outlook ein Add-In für bestimmte Bedingungen aktiviert werden soll, geben Sie Aktivierungsregeln im Add-In-Manifest an. Hierfür stehen zwei  **Rule**-Elemente zur Verfügung:


- [Rule-Element (MailApp complexType) (Add-ins für Office-Schema v1.1)](http://msdn.microsoft.com/de-de/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx), das eine einzelne Regel festlegt
    
- [Rule-Element (RuleCollection complexType) (Add-ins für Office-Schema v1.1)](http://msdn.microsoft.com/de-de/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx), das mehrere Regeln mithilfe logischer Operationen kombiniert
    

(../../images  Das [Rule](http://msdn.microsoft.com/de-de/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)-Element, das Sie zum Angeben einer einzelnen Regel verwenden, gehört dem abstrakten komplexen Typ [Rule](http://msdn.microsoft.com/de-de/library/bcd7a3a7-9cd4-a270-89e0-5386d1c6df01%28Office.15%29.aspx) an. Dieser abstrakte komplexe Typ **Rule** wird durch jede der folgenden Arten von Regeln erweitert. Wenn Sie also eine einzelne Regel in einem Manifest angeben, müssen Sie das [xsi:type](http://www.w3.org/TR/xmlschema-1/)-Attribut verwenden, um eine der folgenden Arten von Regeln näher zu definieren. Die folgende Regel definiert beispielsweise eine [ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)-Regel: `<Rule xsi:type="ItemIs" ItemType="Message" />`Hinweis: Das  **FormType**-Attribut gilt für Aktivierungsregeln im Manifest v1.1, ist jedoch in  **VersionOverrides**nicht definiert. Daher kann es nicht eingesetzt werden, wenn [ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) im **VersionOverrides** -Knoten verwendet wird.

In der folgenden Tabelle werden die verfügbaren Regelntypen aufgeführt. Weitere Informationen finden Sie im Anschluss an die Tabelle und in den unter [Erstellen von Outlook-Add-Ins für Leseformulare](../outlook/read-scenario.md) angegebenen Artikeln.


**In Leseformularen verfügbare Regeltypen**


|**Regelname**|**Anwendbare Formulare**|**Beschreibung**|
|:-----|:-----|:-----|
|[ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)|Lesen, Verfassen, Benutzerdefinierter Bereich|Überprüft, ob das aktuelle Element den angegebenen Typ (Nachricht oder Termin) und (optional) die angegebene Elementnachrichtenklasse aufweist.|
|[ItemHasAttachment](http://msdn.microsoft.com/de-de/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob das ausgewählte Element über einen Anhang verfügt.|
|[ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob das ausgewählte Element über eine oder mehrere allgemein bekannte Entitäten verfügt.Weitere Informationen: [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)|Lesen, Benutzerdefinierter Bereich|Überprüft, ob die E-Mail-Adresse des Absenders, der Betreff und/oder der Textkörper des ausgewählten Elements mit einem regulären Ausdruck übereinstimmt.Weitere Informationen: [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](http://msdn.microsoft.com/de-de/library/926249ab-2d2f-39f5-1d73-fab1c989966f%28Office.15%29.aspx)|Lesen, Verfassen, Benutzerdefinierter Bereich|Kombiniert einen Satz von Regeln, damit Sie komplexere Regeln erstellen können.|

## ItemIs-Regel


Der komplexe Typ  **ItemIs** definiert eine Regel, deren Auswertung **true** ergibt, wenn das aktuelle Element dem Elementtyp und optional der Elementnachrichtenklasse entspricht, sofern diese in der Regel angegeben ist.

Geben Sie einen der folgenden Elementtypen in das Attribut  **ItemType** einer **ItemIs**-Regel ein. Sie können mehr als eine  **ItemIs**Regel in einem Manifest angeben. Der [ItemType](http://msdn.microsoft.com/de-de/library/5a890b98-3d83-77ef-ef03-9b513d35b79f%28Office.15%29.aspx)-SimpleType legt die Outlook-Elementtypen fest, die Outlook-Add-Ins unterstützen.



|**ItemType-Wert**|**Beschreibung**|
|:-----|:-----|
|**Appointment**|Gibt ein Element in einem Outlook-Kalender an. Dies kann eine Einladung sein, ein Besprechungselement, das beantwortet wurde und einen Organisator und Teilnehmer hat, oder ein Termin, der keinen Organisator oder Teilnehmer hat und einfach ein Element im Kalender darstellt.Entspricht der IPM.Appointment-Nachrichtenklassen in Outlook:|
|**Message**|Gibt eines der folgenden, typischerweise im Posteingang empfangenen Elemente an:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Eine E-Mail-Nachricht. Entspricht der Nachrichtenklasse "IPM.Note" in Outlook.</p></li><li><p>Eine Besprechung, Antwort oder ein Abbruch. Entspricht den folgenden Nachrichtenklassen in Outlook:</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|
Das  **FormType**-Attribut wird zum Angeben des Modus (Lesen oder Verfassen) verwendet, in dem das Add-In aktiviert werden sollte.


(../../images  Das [ItemIs](http://msdn.microsoft.com/de-de/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) **FormType**-Attribut ist im Schema v1. 1 und höher, aber nicht in  **VersionOverrides** v1.0 definiert. Fügen Sie nicht das **FormType**-Attribut ein, wenn Sie Regeln für einen benutzerdefinierten Bereich definieren.

Nach der Aktivierung eines Add-Ins können Sie mit der Eigenschaft [mailbox.item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) das aktuell ausgewählte Element in Outlook und mit der Eigenschaft [item.itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) die Art des aktuell ausgewählten Elements abrufen.

Optional können Sie mit dem  **ItemClass**-Attribut die Nachrichtenklasse des Elements angeben, und mit dem  **IncludeSubClasses**-Attribut können Sie angeben, ob die Regel " **true**" sein soll, wenn es sich bei dem Element um eine Unterklasse der angegebenen Klasse handelt.

Weitere Informationen zu Nachrichtenklassen finden Sie unter [Elementtypen und Meldungsklassen](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx).

Das folgende Beispiel ist eine  **ItemIs**-Regel, mit deren Hilfe Benutzer das Add-In in der Outlook-Add-In-Leiste sehen können, wenn der Benutzer eine Nachricht oder einen Termin liest.




```
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

Das folgende Beispiel ist eine  **ItemIs**-Regel, mit deren Hilfe Benutzer das Add-In in der Outlook-Add-In-Leiste sehen können, wenn der Benutzer eine Nachricht liest.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## ItemHasAttachment-Regel


Der komplexe Typ  **ItemHasAttachment** definiert eine Regel, die überprüft, ob das ausgewählte Element eine Anlage enthält.


```XML
<Rule xsi:type="ItemHasAttachment" />
```


## ItemHasKnownEntity-Regel


Bevor ein Element für ein Add-In verfügbar gemacht wird, wird es vom Server untersucht, um festzustellen, ob der Betreff oder der Textkörper Text enthält, bei dem es sich um eine der bekannten Entitäten handeln könnte. Wird eine dieser Entitäten gefunden, wird sie in eine Sammlung bekannter Entitäten verschoben. Diese können Sie mit der  **getEntities**- oder der  **getEntitiesByType**-Methode dieses Elements aufrufen.

Mithilfe des komplexen Typs [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) können Sie eine Regel angeben, mit der Ihr Add-In angezeigt wird, wenn eine Entität des angegebenen Typs in dem Element vorhanden ist. Im **EntityType**-Attribut einer  **ItemHasKnownEntity**-Regel können die folgenden bekannten Entitäten angegeben werden:


-  **Address**
    
-  **Contact**
    
-  **EmailAddress**
    
-  **MeetingSuggestion**
    
-  **PhoneNumber**
    
-  **TaskSuggestion**
    
-  **URL**
    
Diese Entitäten sind im einfachen Typ " [KnownEntityType](http://msdn.microsoft.com/de-de/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx)" als enumerierte Werte definiert.

Optional können Sie einen regulären Ausdruck in das  **RegularExpression**-Attribut integrieren, damit Ihr Add-In nur dann angezeigt wird, wenn eine Entität vorhanden ist, die mit dem regulären Ausdruck übereinstimmt. Um Übereinstimmungen mit regulären Ausrücken abzurufen, die in  **ItemHasKnownEntity**-Regeln angegeben sind, können Sie die Methode  **getRegExMatches** oder **getFilteredEntitiesByName** für das aktuell ausgewählte Outlook-Element verwenden.

Im folgenden Beispiel finden Sie eine Sammlung von  **Rule**-Elementen, mit denen das Add-in angezeigt wird, wenn sich eine der angegebenen bekannten Entitäten in der Nachricht befindet.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

Im folgenden Beispiel finden Sie eine  **ItemHasKnownEntity**-Regel mit einem  **RegularExpression**-Attribut, mit dem das Add-In aktiviert wird, wenn sich eine URL in der Nachricht befindet, die das Wort "contoso" enthält.




```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Weitere Informationen zu Entitäten in Aktivierungsregeln finden Sie unter [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md).


## ItemHasRegularExpressionMatch-Regel


Mithilfe des komplexen Typs  **ItemHasRegularExpressionMatch** wird eine Regel definiert, in der ein regulärer Ausdruck verwendet wird, um die Inhalte der angegebenen Eigenschaft eines Elements abzugleichen. Wird in der angegebenen Eigenschaft des Elements Text gefunden, der mit dem regulären Ausdruck übereinstimmt, wird das Add-In von Outlook aktiviert und auf der Add-In-Leiste angezeigt. Mithilfe der Methode **getRegExMatches** oder **getRegExMatchesByName** des Objekts, das für das aktuell ausgewählte Element steht, können Sie Übereinstimmungen für den angegebenen regulären Ausdruck abrufen.

Im folgenden Beispiel ist ein  **ItemHasRegularExpressionMatch** dargestellt, der das Add-In aktiviert, wenn der Textkörper des ausgewählten Elements das Wort "apple", "banana" oder "coconut" enthält (Groß-/Kleinschreibung wird ignoriert).




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Weitere Informationen zur Verwendung der  **ItemHasRegularExpressionMatch**-Regel finden Sie unter [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).


## RuleCollection-Regel


Mithilfe des komplexen Typs " **RuleCollection**" werden mehrere Regeln in einer einzigen Regel kombiniert. Mit dem  **Mode**-Attribut können Sie angeben, ob die Regeln in der Sammlung mit einem logischen OR oder einem logischen AND kombiniert werden.

Wird ein logisches AND angegeben, muss ein Element mit allen in der Sammlung angegebenen Regeln übereinstimmen, damit das Add-In angezeigt wird. Wird ein logisches OR angegeben, wird das Add-in angezeigt, wenn ein Element mit einer der in der Sammlung angegebenen Regeln übereinstimmt.


(../../images  Das [Rule](http://msdn.microsoft.com/de-de/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx)-Element, mit dem eine Sammlung von Regeln angegeben wird, gehört dem abstrakten komplexen Typ " [Rule](http://msdn.microsoft.com/de-de/library/bcd7a3a7-9cd4-a270-89e0-5386d1c6df01%28Office.15%29.aspx)" an. Der abstrakte komplexe Typ " **Rule**" wird durch den komplexen Typ " **RuleCollection**" erweitert. Wenn Sie also in einem Manifest eine Regelsammlung angeben, müssen Sie das  **xsi:type**-Attribut verwenden, um den komplexen Typ " **RuleCollection**" anzugeben.

Sie können  **RuleCollection**-Regeln kombinieren, um komplexe Regeln zu erstellen. Im folgenden Beispiel wird das Add-In aktiviert, wenn der Benutzer einen Termin oder eine Nachricht anzeigt und der Betreff oder der Textkörper des Elements eine Adresse enthält. 




```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

Mit dem folgenden Beispiel wird das Add-In aktiviert, wenn der Benutzer eine Nachricht verfasst oder wenn der Benutzer einen Termin anzeigt und der Betreff oder der Textkörper des Termins eine Adresse enthalten.




```XML
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## Einschränkungen für Regeln und reguläre Ausdrücke


Für eine zufriedenstellende Erfahrung mit Outlook-Add-Ins sollten Sie die Richtlinien für die Aktivierung und die API-Verwendung befolgen. In der folgenden Tabelle sind die allgemeinen Einschränkungen für reguläre Ausdrücke und Regeln dargestellt, es gibt jedoch spezifische Regeln für unterschiedliche Hosts. Weitere Informationen finden Sie unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md) und [Problembehandlung für die Aktivierung von Outlook-Add-Ins](../../outlook/troubleshoot-outlook-add-in-activation.md).


|
|
|**Add-In-Element**|**Richtlinien**|
|:-----|:-----|
|Manifestgröße|Nicht größer als 32 KB|
|Regeln|Nicht mehr als 15 Regeln.|
|[ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)|Ein Outlook-Rich-Client wendet die Regel auf das erste MB des Elementtexts an, jedoch nicht auf den über diesen Grenzwert hinausgehenden Textkörper.|
|Reguläre Ausdrücke|Für [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)- oder [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regeln für alle Outlook-Hosts gilt:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Sie können nicht mehr als 5 reguläre Ausdrücke in Aktivierungsregeln in einem Outlook-Add-In angeben. Bei Überschreitung dieses Grenzwerts ist die Installation eines Outlook-Add-Ins nicht möglich.</p></li><li><p>Legen Sie die regulären Ausdrücke so fest, dass die von Ihnen erwarteten Ergebnisse vom <span class="keyword">getRegExMatches</span>-Methodenaufruf innerhalb der ersten 50 Übereinstimmungen zurückgegeben wird. </p></li><li><p>Geben Sie Look-Ahead-Assertionen in regulären Ausdrücken an, jedoch keine Look-Behind- (?<=text) und negative Look-Behind-Assertionen (?<!text).</p></li><li><p>Geben Sie reguläre Ausdrücke an, deren Übereinstimmung die folgenden Grenzwerte nicht überschreitet.</p><div class="tableSection"><table border="1" width="50%" cellspacing="2" cellpadding="5" frame="lhs"><tbody><tr><th><p>Grenzwert für eine Übereinstimmung für einen regulären Ausdruck</p></th><th><p>Outlook-Rich-Clients</p></th><th><p>Outlook Web App für Geräte</p></th></tr><tr><td><p>Elementtext ist Nur-Text</p></td><td><p>1,5 KB</p></td><td><p>3 KB</p></td></tr><tr><td><p>Elementtext ist HTML</p></td><td><p>3 KB</p></td><td><p>3 KB</p></td></tr></tbody></table></div></li></ul>|

## Zusätzliche Ressourcen



- [Outlook-Add-Ins](../outlook/outlook-add-ins.md)
    
- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](../outlook/compose-scenario.md)
    
- [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Elementtypen und Meldungsklassen](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
