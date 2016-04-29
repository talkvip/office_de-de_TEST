
# Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten
Nutzen Sie die Funktionen von Exchange Server, um bestimmte Zeichenfolgenmuster zu erkennen, um ein Outlook-Add-In zu aktivieren oder die weitere Verarbeitung dieserZeichenfolgenübereinstimmungen zu erleichtern.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Was sind bekannte Entitäten


Vor dem Senden einer Nachricht oder Besprechungsanfrage analysiert Exchange Server den Inhalt des Elements und identifiziert und stempelt bestimmte Zeichenfolgen im Betreff und im Textkörper, die Exchange bekannten Entitäten ähneln, z. B. E-Mail-Adressen, Telefonnummern und URLs. Nachrichten und Besprechungsanfragen werden von Exchange Server in einem Outlook-Posteingang bereitgestellt, wobei bekannte Identitäten gestempelt werden. Mithilfe der JavaScript-API für Office können Sie diese Zeichenfolgen, die bekannten Entitäten entsprechen, zur weiteren Verarbeitung abrufen. Sie können im Add-in-Manifest mithilfe einer Regel auch eine bekannte Entität festlegen, sodass Outlook Ihr Add-In aktivieren kann, wenn der Benutzer ein Element anzeigt, das Übereinstimmungen mit dieser Entität enthält. Sie können Übereinstimmungen mit der Entität extrahieren und Aktionen daran ausführen. 

Es ist hilfreich, solche Instanzen in einer ausgewählten Nachricht oder einem Termin erkennen oder extrahieren zu können. Sie können beispielsweise eine Umkehrsuche nach Telefonnummern als Outlook-Add-In erstellen. Das Add-in kann Zeichenfolgen im Betreff oder Nachrichtenkörper des Elements, die einer Telefonnummer ähneln, extrahieren, eine Umkehrsuche durchführen und dann den registrierten Eigentümer jeder Telefonnummer anzeigen.

In diesem Thema werden diese bekannten Entitäten vorgestellt und Beispiele für Aktivierungsregeln gegeben, die auf bekannten Entitäten basieren. Außerdem erfahren Sie, wie Entitäten mit und ohne derartige Aktivierungsregeln extrahiert werden.


## Unterstützung für bekannte Entitäten


Exchange Server stempelt bekannte Entitäten in einer Nachricht oder Besprechungsanfrage, nachdem der Absender das Element gesendet hat und bevor Exchange das Element an den Empfänger liefert. Daher werden nur Elemente gestempelt, die von Exchange transportiert wurden, und Outlook kann Add-Ins auf Basis dieser Stempel aktivieren, wenn der Benutzer diese Elemente anzeigt. Wenn der Benutzer allerdings ein Element verfasst oder ein Element im Ordner „Gesendete Objekte" anzeigt, kann Outlook Add-Ins nicht auf Basis bekannter Entitäten aktivieren, da das Element nicht von Exchange transportiert wurde. Entsprechend können Sie bekannte Entitäten nicht in Elementen extrahieren, die verfasst werden oder sich im Ordner „Gesendete Objekte" befinden, da diese Elemente nicht von Exchange transportiert und nicht gestempelt wurden. Weitere Informationen zu den Arten von Elementen, die Aktivierung unterstützen, finden Sie unter [](activation-rules.md#MailAppDefineRules_Activation).

In der folgenden Liste werden die Entitäten aufgeführt, die Exchange Server und Outlook unterstützen und erkennen (daher der Name "bekannte Entitäten") sowie auch den Objekttyp der Instanz einer jeden Entität. Die natürliche Spracherkennung einer Zeichenfolge einer dieser Entitäten basiert auf einem Lernmodell, das mit einer großen Datenmenge verbessert wurde. Deshalb ist die Erkennung nicht deterministisch. Weitere Informationen zu den Erkennungsbedingungen finden Sie unter [Tipps zur Verwendung bekannter Entitäten](#tipps-zur-verwendung-bekannter-entitäten).

 **Tabelle 1. Unterstützte Entitäten und deren Typen**



|**Entitätstyp**|**Bedingung für die Erkennung**|**Objekttyp**|
|:-----|:-----|:-----|
|**Address**|US-Anschrift, wie beispielsweise: 1234 Main Street, Redmond, WA 07722.Im Allgemeinen gilt für die Erkennung einer Adresse, dass sie sich an die Struktur einer US-Postanschrift halten sollte und die meisten Elemente wie Hausnummer, Straßenname, Ort, Bundesstaat und PLZ vorhanden sein sollten. Die Adresse kann in einer oder mehreren Zeilen angegeben werden.|**String**-Objekt (JavaScript)|
|**Contact**|Ein Verweis auf die Informationen einer Person, die in natürlicher Sprache erkannt werden.Die Erkennung eines Kontakts hängt vom Kontext ab. Beispielsweise eine Signatur am Ende einer Nachricht, oder der Name einer Person in der Nähe der folgenden Informationen: Telefonnummer, Adresse, E-Mail-Adresse und URL.|[Contact](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt|
|**EmailAddress**|SMTP-E-Mail-Adressen|**String**-Objekt (JavaScript)|
|**MeetingSuggestion**|Ein Verweis auf ein Ereignis oder eine Besprechung. Beispielsweise würde Exchange 2013 den folgenden Text als Besprechungsvorschlag erkennen:  _Treffen wir uns morgen zum Mittagessen._|[MeetingSuggestion](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt|
|**PhoneNumber**|US-Telefonnummern, wie beispielsweise:  _(235) 555-0110_|[PhoneNumber](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt|
|**TaskSuggestion**|Sätze in einer E-Mail, die zu Aktionen auffordern. Beispiel:  _Bitte aktualisieren Sie das Arbeitsblatt._|[TaskSuggestion](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt|
|**Url**|Eine Webadresse, die den neuen Speicherort des Netzwerks sowie die ID einer Webressource explizit angibt. Exchange Server fordert nicht das Zugriffsprotokoll in der Webadresse an und erkennt auch nicht die URLs, die in Linktext eingebettet sind, als Instanzen der  **Url**-Entität. Exchange Server kann die folgenden Beispiel zuordnen: _www.youtube.com/user/officevideos_ _http://www.youtube.com/user/officevideos_|**String**-Objekt (JavaScript)|
In Abbildung 1 wird gezeigt, auf welche Weise bekannte Entitäten für Add-Ins von Exchange Server und Outlook unterstützt werden und welche Aktionen Add-Ins mit bekannten Entitäten ausführen können. Weitere Informationen zur Verwendung dieser Entitäten finden Sie unter [Abrufen von Entitäten in Ihrem Add-In](#abrufen-von-entitäten-in-ihrem-add-in) und [Aktivieren eines Add-Ins basierend auf dem Vorhandensein einer Entität](#aktivieren-eines-add-ins-basierend-auf-dem-vorhandensein-einer-entität).


**Abbildung 1: Unterstützung bekannter Entitäten durch Exchange Server, Outlook und Add-Ins**

![Unterstützung und Verwendung bekannter Entitäten in der Mailanwendung](../../images/mod_off15_mailapp_wellknownentities_curvedlines.png)


## Berechtigungen zum Extrahieren von Entitäten


Fordern Sie unbedingt im Add-In-Manifest die entsprechenden Berechtigungen an, um Entitäten in Ihrem JavaScript-Code extrahieren zu können oder damit Ihr Add-In basierend auf dem Vorhandensein bestimmter bekannter Entitäten aktiviert wird.

Durch die Angabe der Standardberechtigung „eingeschränkt" kann Ihr Add-In die Entität  **Address**,  **MeetingSuggestion** oder **TaskSuggestion** extrahieren. Zum Extrahieren anderer Entitäten geben Sie die Berechtigung „Element lesen" oder „Postfach lesen/schreiben" an. Um dies im Manifest durchzuführen, verwenden Sie das Element [Permissions](http://msdn.microsoft.com/de-de/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx), und geben Sie die entsprechende Berechtigung –  **Restricted**,  **ReadItem**,  **ReadWriteItem**, or  **ReadWriteMailbox** – wie im folgenden Beispiel gezeigt an:




```XML
<Permissions>ReadItem</Permissions>
```


## Abrufen von Entitäten in Ihrem Add-In


Solange der Betreff oder der Textkörper des vom Benutzer angezeigten Elements Zeichenfolgen enthält, die Exchange und Outlook als bekannte Entitäten erkennen können, sind diese Instanzen für Add-Ins verfügbar. Dies gilt sogar dann, wenn ein Mail-Add-In nicht auf Basis bekannter Entitäten aktiviert wurde. Mit der entsprechenden Berechtigung können Sie die  **getEntities**- oder die  **getEntitiesByType**-Methode verwenden, um bekannte Entitäten abzurufen, die in der aktuellen Nachricht oder dem Termin enthalten sind. Die  **getEntities**-Methode gibt ein Array von [Entities](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekten zurück, das alle bekannten Entitäten in dem Element enthält. Wenn Sie an einem bestimmten Typ von Entitäten interessiert sind, verwenden Sie die  **getEntitiesByType**-Methode, die ein Array mit den von Ihnen gewünschten Entitäten zurückgibt. Die [EntityType](http://dev.outlook.com/reference/add-ins/Office.MailboxEnums.html%28Office.15%29.md)-Enumeration stellt alle Arten bekannter Entitäten dar, die Sie extrahieren können.

Nach dem Aufruf von  **getEntities** können Sie anschließend mithilfe der entsprechenden Eigenschaft des **Entities**-Objekts ein Array von Instanzen eines Entitätstyps abrufen. In Abhängigkeit vom Entitätstyp können die Instanzen im Array nur Zeichenfolgen sein oder aber bestimmten Objekten zugeordnet sein. Wie in dem Beispiel in Abbildung 1 gezeigt, greifen Sie zum Abrufen der Adressen in dem Element auf das von  `getEntities().addresses[]` zurückgegebene Array zu. Die **Entities.addresses**-Eigenschaft gibt ein Array von Zeichenfolgen zurück, die Outlook als Postanschriften erkennt. Ebenso gibt die  **Entities.contacts**-Eigenschaft ein Array von  **Contact**-Objekten zurück, die Outlook als Kontaktinformationen erkennt. In Tabelle 1 ist der Objekttyp einer Instanz jeder unterstützten Entität aufgeführt.

Im folgenden Beispiel wird gezeigt, wie Sie in einer Nachricht gefundene Adressen abrufen.




```
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities &amp;&amp; null != entities.addresses &amp;&amp; undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## Aktivieren eines Add-Ins basierend auf dem Vorhandensein einer Entität


Eine weitere Möglichkeit für die Verwendung bekannter Entitäten besteht darin, Outlook die Aktivierung Ihres Add-Ins basierend auf dem Vorhandensein von einem oder mehreren Entitätstypen im Betreff oder Textkörper des aktuell angezeigten Elements zu überlassen. Geben Sie hierfür eine  **ItemHasKnownEntity**-Regel im Add-In-Manifest an. Der einfache [KnownEntityType](http://msdn.microsoft.com/de-de/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx)-Typ stellt die verschiedenen Typen bekannter Entitäten dar, die von  **ItemHasKnownEntity**-Regeln unterstützt werden. Nachdem Ihr Add-In aktiviert wurde, können Sie auch die Instanzen derartiger Entitäten für Ihre Zwecke abrufen. Dies wird im vorherigen Abschnitt [Abrufen von Entitäten in Ihrem Add-In](#abrufen-von-entitäten-in-ihrem-add-in) näher erläutert.

Optional können Sie einen regulären Ausdruck in einer  **ItemHasKnownEntity**-Regel anwenden, damit Instanzen einer Entität weiter gefiltert werden und Outlook ein Add-In nur für einen Teil der Entitätsinstanzen aktiviert. Beispielsweise können Sie einen Filter für die Postanschriftenentität in einer Nachricht angeben, die eine Postleitzahl für den Bundesstaat Washington enthält und mit "98" beginnt. Zum Anwenden eines Filters auf die Entitätsinstanzen verwenden Sie die Attribute  **RegExFilter** und **FilterName** im [Rule](http://msdn.microsoft.com/de-de/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)-Element des [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Typs. 

Ähnlich wie bei anderen Aktivierungsregeln können Sie mehrere Regeln angeben, um eine Regelsammlung für Ihr Add-In zu bilden. Im folgenden Beispiel wird eine Operation vom Typ AND auf 2 Regeln angewendet: eine  **ItemIs**-Regel und eine  **ItemHasKnownEntity**-Regel. Diese Regelsammlung aktiviert das Add-In immer dann, wenn das aktuelle Element eine Nachricht ist und Outlook eine Adresse im Betreff oder im Textkörper des Elements erkennt.




```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

Im folgenden Beispiel wird mithilfe von  **getEntitiesByType** des aktuellen Elements die Variable `addresses` auf die Ergebnisse der vorherigen Regelsammlung festgelegt.




```
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

Die folgende  **ItemHasKnownEntity**-Regelbeispiel aktiviert das Add-In, wenn es eine URL im Betreff oder Textkörper des aktuellen Elements gibt und die URL die Zeichenfolge „youtube" ungeachtet von Groß-/Kleinschreibung enthält.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

Im folgenden Beispiel wird mithilfe von  **getFilteredEntitiesByName(name)** des aktuellen Elements die Variable `videos` festgelegt, um ein Array der Ergebnisse abzurufen, die mit dem regulären Ausdruck in der vorherigen **ItemHasKnownEntity**-Regel übereinstimmen.




```
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## Tipps zur Verwendung bekannter Entitäten


Es gibt einige wenige Aspekte und Beschränkungen, die Sie beachten sollten, wenn Sie bekannte Entitäten in Ihrem Add-In verwenden. Die folgenden Angaben gelten, solange Ihr Mail-Add-in aktiviert ist, wenn der Benutzer ein Element liest, das Übereinstimmungen mit bekannten Entitäten enthält, unabhängig davon, ob Sie eine  **ItemHasKnownEntity**-Regel verwenden:


1. Sie können Zeichenfolgen, bei denen es sich um bekannte Entitäten handelt, nur dann extrahieren, wenn die Zeichenfolgen in Englisch vorliegen.
    
2. Sie können bekannte Entitäten aus den ersten 2.000 Zeichen im Elementtext extrahieren, aber nicht über diesen Grenzwert hinaus. Diese Größenbeschränkung hilft dabei, den Anspruch an Funktionalität und Leistung ins Gleichgewicht zu bringen, sodass sich Exchange Server und Outlook nicht beim Analysieren und Ermitteln von Instanzen bekannter Entitäten in umfangreichen Nachrichten und Terminen verzetteln. Beachten Sie, dass diese Beschränkung unabhängig davon besteht, ob das Add-In eine  **ItemHasKnownEntity**-Regel angibt. Wenn das Mail-Add-in eine solche Regel verwendet, müssen Sie auch das Verarbeitungslimit für Regeln in Element 2 unten für Outlook-Rich-Clients beachten.
    
3. Sie können Entitäten aus Terminen extrahieren, bei denen es sich um Besprechungen handelt, die von einer anderen Person als dem Besitzer des Postfachs organisiert wurden. Sie können keine Entitäten aus Kalenderelementen extrahieren, bei denen es sich nicht um Besprechungen oder aber um vom Besitzer des Postfachs organisierte Besprechungen handelt.
    
4. Sie können Entitäten vom Typ  **MeetingSuggestion** nur aus Nachrichten, nicht aber aus Terminen extrahieren.
    
5. Sie können URLs extrahieren, die explizit im Elementtextkörper vorhanden sind, jedoch keine URLs, die in Hyperlinktext im HTML-Elementtextköper eingebettet sind. Erwägen Sie die Verwendung einer  **ItemHasRegularExpressionMatch**-Regel, um nicht sowohl explizite als auch eingebettete URLs zu erhalten. Legen Sie  **BodyAsHTML** als _PropertyName_ fest, und einen regulären Ausdruck, der mit den URLs übereinstimmt, als _RegExValue_.
    
6. Entitäten können nicht aus Elementen im Ordner  **"Gesendete Elemente"** extrahiert werden.
    
Außerdem gilt Folgendes, wenn Sie eine [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regel verwenden, das sich möglicherweise auch auf die Szenarien auswirkt, in denen Sie andernfalls erwarten, dass Ihr Add-In aktiviert wird:


1. Bei Verwendung der  **ItemHasKnownEntity**-Regel gleicht Outlook Entitätszeichenfolgen unabhängig von dem im Manifest angegebenen Standardgebietsschema nur auf Englisch ab.
    
2. Wenn Ihr Add-In auf einem Outlook-Rich-Client ausgeführt wird, wendet Outlook die  **ItemHasKnownEntity**-Regel auf das erste MB des Elementtexts an, jedoch nicht auf den über diesen Grenzwert hinausgehenden Textkörper.
    
3. Sie können keine  **ItemHasKnownEntity**-Regel verwenden, um ein Add-In für Elemente im Ordner „Gesendete Elemente" zu aktivieren.
    

## Weitere Ressourcen



- [Erstellen von Outlook-Add-Ins für Leseformulare](c8680c4e-3068-4e13-8537-2a96c614193e.md)
    
- [Extrahieren von Entitätszeichenfolgen aus einem Outlook-Element](50b70957-78a7-47ba-93bd-f4d7b6ff50b7.md)
    
- [Aktivierungsregeln für Outlook-Add-Ins](b3fd6d69-b968-461d-a40e-6063f4febfe6.md)
    
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](93504f92-896f-4c80-9205-ba0b125f4290.md)
    
- [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](5bca69f2-b287-4e19-8f0f-78d896b2a3d3.md)
    
