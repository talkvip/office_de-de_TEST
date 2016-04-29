
# Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins
Erfahren Sie, wie Sie reguläre Ausdrücke verwenden, um die Kriterien dafür anzugeben, wann ein Outlook-Add-In in einem Leseszenario relevant ist und Outlook das Add-In in der UI anzeigen soll.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Unterstützung regulärer Ausdrücke
<a name="MailAppRegExpressions_Support"> </a>

Sie können reguläre Ausdrucksregeln so festlegen, dass in Leseszenarios ein Outlook-Add-In aktiviert wird – wenn der Benutzer im Lesebereich oder Inspektor eine Nachricht oder einen Termin anzeigt, bewertet Outlook Regeln mit regulären Ausdrücken, um zu ermitteln, ob Ihr Add-In aktiviert werden sollte. Outlook bewertet diese Regeln nicht, wenn der Benutzer ein Element verfasst. Es gibt auch andere Szenarios, in denen Outlook Add-Ins nicht aktiviert. Dies gilt z. B. für Elemente, die durch das Information Rights Management (IRM) geschützt sind oder sich im Junk-E-Mail-Ordner befinden. Weitere Informationen finden Sie unter [](../outlook/manifests/activation-rules.md#MailAppDefineRules_Activation).

Sie können einen regulären Ausdruck im XML-Manifest des Add-ins als Teil einer [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)- oder [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regel angeben. Outlook wertet reguläre Ausdrücke auf Basis der Regeln für den JavaScript-Interpreter aus, der vom Browser auf dem Clientcomputer verwendet wird. Outlook unterstützt dieselbe Liste von Sonderzeichen, die alle XML-Prozessoren ebenfalls unterstützen. Die folgende Tabelle enthält diese Sonderzeichen. Sie können diese Zeichen in einem regulären Ausdruck verwenden, indem Sie die Escapesequenz für das entsprechende Zeichen eingeben (siehe folgende Tabelle).



|**Zeichen**|**Beschreibung**|**Zu verwendende Escapesequenz**|
|:-----|:-----|:-----|
|"|Doppeltes Anführungszeichen|&amp;quot;|
|&amp;|Kaufmännisches Und-Zeichen|&amp;amp;|
|‘|Apostroph|&amp;apos;|
|<|Kleiner als-Zeichen|&amp;lt;|
|>|Größer als-Zeichen|&amp;gt;|

## "ItemHasRegularExpressionMatch"-Regel
<a name="MailAppRegExpressions_ItemHasRegularExpressionMatch"> </a>

Eine  **ItemHasRegularExpressionMatch**-Regel eignet sich zum Steuern der Aktivierung eines Add-Ins basierend auf bestimmten Werten einer unterstützten Eigenschaft. Die  **ItemHasRegularExpressionMatch**-Regel hat die folgenden Attribute.



|**Name des Attributs**|**Beschreibung**|
|:-----|:-----|
|**RegExName**|Gibt den Namen des regulären Ausdrucks an, damit Sie im Code Ihres Add-ins auf den Ausdruck verweisen können.|
|**RegExValue**|Gibt den regulären Ausdruck an, der ausgewertet wird, um zu bestimmen, ob das Add-In angezeigt werden soll.|
|**PropertyName**|Gibt den Namen der Eigenschaft an, für die der reguläre Ausdruck ausgewertet wird. Die zulässigen Werte sind  **BodyAsHTML**,  **BodyAsPlaintext**,  **SenderSMTPAddress** und **Subject**. Wenn Sie  **BodyAsHTML** angeben, wendet Outlook den regulären Ausdruck nur an, wenn der Elementtext HTML ist. Andernfalls gibt Outlook keine Übereinstimmungen für diesen regulären Ausdruck zurück. Da Termine stets im RTF-Format gespeichert werden, stimmt ein regulärer Ausdruck, der **BodyAsHTML** angibt, nicht mit Zeichenfolgen im Textköper von Terminelementen überein.Wenn Sie  **BodyAsPlaintext** angeben, wendet Outlook stets den regulären Ausdruck auf den Textköper des Elements an.|
|**IgnoreCase**|Gibt an, ob die Schreibung ignoriert werden soll, wenn ein Abgleich mit dem von  **RegExName** angegebenen regulären Ausdruck erfolgt.|

### Bewährte Methoden für das Verwenden regulärer Ausdrücke in Regeln

Beachten Sie beim Verwenden regulärer Ausdrücke unbedingt Folgendes:


- Wenn Sie eine  **ItemHasRegularExpressionMatch**-Regel auf den Textkörper eines Elements anwenden, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückgeben. Das Verwenden eines regulären Ausdrucks wie  `.*`, um zu versuchen, den gesamten Textkörper eines Elements abzurufen, liefert nicht immer die erwarteten Ergebnisse.
    
- Der in einem Browser zurückgegebene Nur-Text-Körper kann in einem anderen geringfügig anders sein. Wenn Sie eine [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regel mit  **BodyAsPlaintext** als **PropertyName**-Attribut verwenden, testen Sie Ihren regulären Ausdruck in allen Browsern, die Ihr Add-in unterstützt.
    
    Da unterschiedliche Browser den Textkörper eines ausgewählten Elements auf verschiedene Weisen abrufen, müssen Sie sicherstellen, dass Ihr regulärer Ausdruck die geringfügigen Unterschiede unterstützt, die als Teil des Textkörpers zurückgegeben werden können. Beispielsweise verwenden einige Browser wie Internet Explorer 9 die  **innerText**-Eigenschaft des DOM, während andere Browser wie Firefox die  **.textContent()**-Methode zum Abrufen des Textkörpers eines Elements nutzen. Des Weiteren geben verschiedene Browser Zeilenumbrüche möglicherweise anders zurück: In Internet Explorer ist "\r\n" ein Zeilenumbruch, in Firefox und Chrome "\n". Weitere Informationen finden Sie unter [W3C-DOM-Kompatibilität – HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07).
    
- Der HTML-Textkörper eines Elements ist bei einem Outlook-Richt-Client und Outlook Web App oder OWA für mobile Geräte leicht unterschiedlich. Definieren Sie reguläre Ausdrücke sorgfältig. Sehen Sie sich z. B. den folgenden regulären Ausdruck an, der in einer  **ItemHasRegularExpressionMatch**-Regel mit  **BodyAsHTML** als Wert des **PropertyName**-Attributs verwendet wird:
    
  ```
  http.*\.contoso\.com
  ```


    Eine Regel mit diesem regulären Ausdruck entspricht der Zeichenfolge "http-equiv="Content-Type", die im HTML-Textkörper eines Elements in einem Outlook-Rich-Client als Teil des folgenden  **META**-Tags vorhanden ist:
    


  ```HTML
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
  ```


    In Outlook Web App und OWA für mobile Geräte gibt dieselbe Regel diese Übereinstimmung nicht zurück, da der HTML-Textkörper in diesen Hosts nicht das  **META**-Tag enthält. Dies kann sich darauf auswirken, ob das Add-In für die verschiedenen Outlook-Clients ordnungsgemäß geöffnet wird. Verwenden Sie bei diesem Beispiel stattdessen den folgenden regulären Ausdruck:
    


  ```
  http://.*\.contoso\.com/
  ```

- Je nach Hostanwendung, Gerätetyp oder Eigenschaft, auf die/den ein regulärer Ausdruck angewendet wird, gibt es andere bewährte Methoden und Einschränkungen für jeden der Hosts, die Sie kennen sollten, wenn Sie reguläre Ausdrücke als Aktivierungsregeln entwerfen. Weitere Informationen finden Sie unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](e0c9e3d0-517e-4333-b8bd-e169c51a07f6.md).
    

### Beispiele

Die folgende  **ItemHasRegularExpressionMatch**-Regel aktiviert das Add-In, wenn die SMTP-E-Mail-Adresse des Absenders „@contoso" entspricht (unabhängig von Groß-/Kleinschreibung).


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

Es folgt eine andere Möglichkeit, denselben regulären Ausdruck mithilfe des  **IgnoreCase**-Attributs anzugeben.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

Die folgende  **ItemHasRegularExpressionMatch**-Regel aktiviert das Add-In, wenn der Textkörper des aktuellen Elements ein Aktiensymbol enthält.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## "ItemHasKnownEntity"-Regel
<a name="MailAppRegExpressions_ItemHasKnownEntity"> </a>

Eine  **ItemHasKnownEntity**-Regel aktiviert ein Add-In basierend auf dem Vorhandensein einer Entität im Betreff oder Textkörper des ausgewählten Elements. Der [KnownEntityType](http://msdn.microsoft.com/de-de/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx)-Typ definiert die unterstützten Entitäten. Das Anwenden eines regulären Ausdrucks auf eine  **ItemHasKnownEntity**-Regel bewirkt die bequeme Aktivierung basierend auf einer Untermenge von Werten für eine Entität (z. B. einer bestimmten Gruppe von URLs oder von Telefonnummern mit einer bestimmten Vorwahl). 


 **Hinweis**  Outlook kann ungeachtet des in der Manifestdatei angegebenen Standardgebietsschemas nur Entitätszeichenfolgen in Englisch extrahieren. Der Entitätstyp  **MeetingSuggestion** wird nur für Nachrichten, nicht aber für Termine unterstützt.Sie können keine Entitäten aus Elementen im Ordner  **Gesendete Objekte** extrahieren. Sie können auch nicht die[ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regel verwenden, um ein Add-In für Elemente im Ordner  **Gesendete Objekte** zu aktivieren.

Die  **ItemHasKnownEntity**-Regel unterstützt die Attribute in der folgenden Tabelle. Wenngleich das Angeben eines regulären Ausdrucks in einer  **ItemHasKnownEntity**-Regel optional ist, müssen Sie beim Auswählen eines regulären Ausdrucks als Entitätsfilter die beiden Attribute  **RegExFilter** und **FilterName** angeben.



|**Name des Attributs**|**Beschreibung**|
|:-----|:-----|
|**EntityType**|Gibt den Typ der Entität an, der gefunden werden muss, damit die Regel mit  **true** ausgewertet wird. Verwenden Sie mehrere Regeln, um mehrere Entitätstypen anzugeben.|
|**RegExFilter**|Gibt einen regulären Ausdruck an, der Vorkommen der von  **EntityType** angegebenen Entität weiter filtert.|
|**FilterName**|Gibt den Namen des regulären Ausdrucks an, der von  **RegExFilter** angegeben wird, damit anschließend auf diesen in Code verwiesen werden kann.|
|**IgnoreCase**|Gibt an, ob die Schreibung ignoriert werden soll, wenn ein Abgleich mit dem von  **RegExFilter** angegebenen regulären Ausdruck erfolgt.|

### Beispiele

Die folgende  **ItemHasKnownEntity**-Regel aktiviert das Add-In, wenn es eine URL im Betreff oder Textkörper des aktuellen Elements gibt und die URL die Zeichenfolge „youtube" ungeachtet von Groß-/Kleinschreibung enthält.


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## Verwenden der Ergebnisse regulärer Ausdrücke in Code
<a name="MailAppRegExpressions_InCode"> </a>

Sie können Übereinstimmungen mit einem regulären Ausdruck erhalten, indem Sie die folgenden Methoden auf das aktuelle Elements anwenden:


- [getRegExMatches](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) gibt Übereinstimmungen im aktuellen Element für alle regulären Ausdrücke zurück, die in den Regeln **ItemHasRegularExpressionMatch** und **ItemHasKnownEntity** des Add-Ins angegeben sind.
    
- [getRegExMatchesByName](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) gibt Übereinstimmungen im aktuellen Element für den angegebenen regulären Ausdruck zurück, der in einer **ItemHasRegularExpressionMatch**-Regel des Add-Ins angegeben ist.
    
- [getFilteredEntitiesByName](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) gibt gesamte Instanzen von Entitäten mit Übereinstimmungen mit dem angegebenen regulären Ausdruck zurück, der in einer **ItemHasKnownEntity**-Regel des Add-Ins angegeben ist.
    
Wenn die regulären Ausdrücke ausgewertet werden, erfolgt die Rückgabe der Übereinstimmungen an Ihr Add-In in einem Arrayobjekt. Bei  **getRegExMatches** hat dieses Objekt als ID den Namen des regulären Ausdrucks.


 **Hinweis**  Ein Outlook-Rich-Client gibt Übereinstimmungen im Array in keiner besonderen Reihenfolge zurück. Außerdem sollten Sie nicht annehmen, dass der Outlook-Rich-Client Übereinstimmungen in derselben Reihenfolge in diesem Array zurückgibt wie Outlook Web App oder OWA für mobile Geräte, auch wenn Sie dasselbe Add-In auf diesen Clients für dasselbe Element in demselben Postfach ausführen. Informationen zu weiteren Unterschieden bei der Verarbeitung regulärer Ausdrücke zwischen einem Outlook-Rich-Client und Outlook Web App oder OWA für mobile Geräte finden Sie unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](e0c9e3d0-517e-4333-b8bd-e169c51a07f6.md).


### Beispiele

Es folgt ein Beispiel einer Regelsammlung, die die  **ItemHasRegularExpressionMatch**-Regel mit dem regulären Ausdruck  `videoURL` enthält.


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

Im folgenden Beispiel wird  **getRegExMatches** des aktuellen Elements verwendet, um die Variable `videos` auf das Ergebnis der vorherigen **ItemHasRegularExpressionMatch**-Regel festzulegen.




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Mehrere Übereinstimmungen werden als Arrayelemente in diesem Objekt gespeichert. Das folgende Codebeispiel veranschaulicht das Durchlaufen der Übereinstimmungen des regulären Ausdrucks  `reg1` zum Erstellen einer Zeichenfolge, die als HTML angezeigt wird.




```
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

Es folgt ein Beispiel der  **ItemHasKnownEntity**-Regel, die die Entität  **MeetingSuggestion** und den regulären Ausdruck `CampSuggestion` angibt. Outlook aktiviert das Add-In, wenn erkannt wird, dass das aktuell ausgewählte Element einen Besprechungsvorschlag enthält und der Betreff oder Textkörper den Begriff „WonderCamp" enthält.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

Im folgenden Codebeispiel wird  **getFilteredEntitiesByName(name)** des aktuellen Elements verwendet, um die Variable `suggestions` so festzulegen, dass ein Array erkannter Besprechungsvorschläge für die vorherige **ItemHasKnownEntity**-Regel abgerufen wird.




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## Zusätzliche Ressourcen
<a name="MailAppRegExpressions_AdditionalResources"> </a>


- [Erstellen von Outlook-Add-Ins für Leseformulare](c8680c4e-3068-4e13-8537-2a96c614193e.md)
    
- [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md)
    
- [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](e0c9e3d0-517e-4333-b8bd-e169c51a07f6.md)
    
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Bewährte Methoden für reguläre Ausdrücke in .NET Framework](http://msdn.microsoft.com/de-de/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
