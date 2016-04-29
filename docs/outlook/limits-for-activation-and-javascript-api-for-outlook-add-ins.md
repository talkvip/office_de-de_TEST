
# Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins
In diesem Thema werden die Grenzwerte für das Verwenden von Aktivierungsregeln und JavaScript-API für Office zum Abrufen oder Festlegen von Daten sowie zum Durchführen asynchroner Aufrufe in Outlook-Add-Ins erläutert.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Um für Benutzer von Outlook-Add-Ins eine höchstmögliche Benutzerfreundlichkeit zu gewährleisten, sollten Sie bestimmte Richtlinien in Bezug auf Aktivierung und API-Verwendung kennen und Ihre Outlook-Add-Ins so implementieren, dass diese Grenzwerte eingehalten werden. Diese Grenzwerte bestehen, damit ein einzelnes Outlook-Add-In Exchange Server oder Outlook nicht veranlassen kann, eine ungewöhnlich lange Zeitspanne für das Verarbeiten ihrer Aktivierungsregeln oder für Aufrufe der JavaScript-API für Office aufzuwenden, was zulasten der allgemeinen Benutzerfreundlichkeit bei Outlook und anderen Outlook-Add-Ins ginge. Diese Grenzwerte gelten für das Entwerfen von Aktivierungsregeln im Add-in-Manifest und für die Verwendung von benutzerdefinierten Eigenschaften, Roamingeinstellungen, Empfängern, Exchange-Webdienste-Anforderungen und -Antworten und asynchronen Aufrufen 

(../../images  Falls Ihr Outlook-Add-in im Outlook-Rich-Client ausgeführt wird, muss das Add-in im Rahmen bestimmter Grenzwerte für die Nutzung von Laufzeitressourcen funktionieren. 


## Grenzwerte für Aktivierungsregeln


Halten Sie die folgenden Richtlinien beim Entwerfen von Aktivierungsregeln für Outlook-Add-Ins ein:


- Beschränken Sie die Größe des Manifests auf 256 KB. Bei Überschreitung dieses Grenzwerts ist die Installation des Outlook-Add-ins für ein Exchange-Postfach nicht möglich.
    
- Legen Sie bis zu 15 Aktivierungsregeln für das Outlook-Add-Ins fest. Bei Überschreitung dieses Grenzwerts ist die Installation des Outlook-Add-Ins nicht möglich.
    
- Wenn Sie eine [ItemHasKnownEntity](http://msdn.microsoft.com/de-de/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)-Regel auf den Textkörper des ausgewählten Elements anwenden, können Sie erwarten, dass ein Outlook-Rich-Client die Regel nur auf die ersten 1 MB des Textkörpers anwendet und nicht auf den restlichen Textkörper. Wären nur nach den ersten 1 MB des Textkörpers Treffer vorhanden, würde Ihr Outlook-Add-In nicht aktiviert werden. Wenn Sie glauben, dass dieses Szenario wahrscheinlich ist, sollten Sie die Aktivierungsbedingungen ändern.
    
- Wenn Sie reguläre Ausdrücke in  **ItemHasKnownEntity** oder [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/de-de/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regeln verwenden, folgen Sie diesen Richtlinien, und achten Sie auf die in den Tabellen 1, 2 und 3 beschriebenen Supporteinschränkungen eines beliebigen Outlook-Hosts, die je nach Host variieren:
    
      - Sie können nur bis zu 5 reguläre Ausdrücke in Aktivierungsregeln in einem Outlook-Add-In angeben. Bei Überschreitung dieses Grenzwerts ist die Installation eines Outlook-Add-Ins nicht möglich.
    
  - Legen Sie die regulären Ausdrücke so fest, dass die von Ihnen erwarteten Ergebnisse vom  **getRegExMatches**-Methodenaufruf innerhalb der ersten 50 Übereinstimmungen zurückgegeben wird. 
    
  - Kann Look-Ahead-Assertionen in regulären Ausdrücken angeben, jedoch keine Look-Behind- (?<=text) und negative Look-Behind-Assertionen (?<!text).
    

    In Tabelle 1 werden die Grenzwerte für die Unterstützung für reguläre Ausdrücke und die diesbezüglichen Unterschiede zwischen einem Outlook-Rich-Client und Outlook Web App oder OWA für mobile Geräte erläutert. Die Unterstützung ist nicht abhängig von spezifischen Gerätetypen und Hauptteilen von Elementen.
    

    **Tabelle 1. Allgemeine Unterschiede bei der Unterstützung für reguläre Ausdrücke**


|**Outlook-Rich-Client**|**Outlook Web App oder OWA für mobile Geräte**|
|:-----|:-----|
|Verwendet das C++-Modul für reguläre Ausdrücke, das als Teil der standardmäßigen Vorlagenbibliothek von Visual Studio bereitgestellt wird. Dieses Modul erfüllt die ECMAScript 5-Normen. |Verwendet eine Auswertung für reguläre Ausdrücke, die Teil von JavaScript ist, über den Browser bereitgestellt wird und eine Obermenge von ECMAScript 5 darstellt.|
|Erwarten Sie, dass ein regulärer Ausdruck mit einer benutzerdefinierten Zeichenklasse auf Grundlage einer vordefinierten Zeichenklasse in einem Outlook-Rich-Client möglicherweise andere Ergebnisse zurückgibt als in Outlook Web App oder OWA für mobile Geräte. Ursache hierfür ist, dass der Outlook-Rich-Client ein anderes RegEx-Modul verwendet. Der reguläre Ausdruck "[\s\S]{0,100}" stimmt zum Beispiel mit jeder Anzahl von 0 bis 100 von einzelnen Zeichen überein, bei der es sich um ein Leerzeichen oder um ein Nicht-Leerzeichen handelt. Von diesem regulären Ausdruck werden in einem Outlook-Rich-Client andere Ergebnisse zurückgegeben als in Outlook Web App und OWA für mobile Geräte. Schreiben Sie, um dieses Problem zu umgehen, den regulären Ausdruck folgendermaßen um: "(\s|\S){0,100}". Dieser reguläre Ausdruck stimmt mit jeder Anzahl von 0 bis 100 von Leerzeichen und Nicht-Leerzeichen überein.Testen Sie jeden RegEx-Ausdruck sorgfältig auf dem Outlook-Host, und schreiben Sie RegEx-Ausdrücke um, die dazu führen, dass abweichende Ergebnisse zurückgegeben werden. |Testen Sie jeden RegEx-Ausdruck sorgfältig auf dem Outlook-Host, und schreiben Sie RegEx-Ausdrücke um, die dazu führen, dass abweichende Ergebnisse zurückgegeben werden. |
|Begrenzt die Überprüfung aller regulären Ausdrücke für ein Outlook-Add-In standardmäßig auf 1 Sekunde. Wird dieser Grenzwert überschritten, wird maximal 3 Mal eine erneute Überprüfung durchgeführt. Wird der Grenzwert für die erneute Überprüfung überschritten, deaktiviert ein Outlook-Rich-Client die Ausführung des Outlook-Add-Ins für dasselbe Postfach auf allen Outlook-Hosts.Mithilfe der Registrierungsschlüssel " **OutlookActivationAlertThreshold**" und " **OutlookActivationManagerRetryLimit**" können Administratoren diese Grenzwerte für die Überprüfung außer Kraft setzen.|Unterstützt nicht dieselben Einstellungen für die Ressourcenüberwachung oder die Registrierung wie ein Outlook-Rich-Client. Outlook-Add-Ins mit regulären Ausdrücken, die auf einem Outlook-Rich-Client zu lange für die Überprüfung brauchen, werden auf allen Outlook-Hosts für dasselbe Postfach deaktiviert.|

    In Tabelle 2 werden die Grenzwerte und Unterschiede hinsichtlich des Teils des Textkörpers des Elements erläutert, auf den Outlook einen regulären Ausdruck anwenden. Einige dieser Grenzwerte hängen von der Art des Geräts und vom Textkörper ab, wenn der reguläre Ausdruck auf den Textkörper des Elements angewendet wird.
    

    **Tabelle 2. Grenzwerte für die Größe des überprüften Hauptteils des Elements**


||**Outlook-Rich-Client**|**Outlook Web App, OWA für mobile Geräte,OWA für iPad oder OWA für iPhone**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|Gerätegrößen|Alle unterstützten Geräte|Android Smartphones, iPad oder iPhone|Alle unterstützten Geräte außer Android Smartphones, iPad und iPhone|
|Nur-Text-Hauptteil des Elements|Wendet den regulären Ausdruck auf das erste MB der Daten im Hauptteil an, nicht jedoch auf den Rest des Hauptteils, der diesen Grenzwert überschreitet.|Aktiviert das Add-in nur, wenn der Hauptteil < 16.000 Zeichen groß ist.|Aktiviert das Add-in nur, wenn der Hauptteil < 500.000 Zeichen groß ist.|
|Hauptteil des HTML-Elements|Wendet den regulären Ausdruck auf die ersten 512 kB der Daten im Hauptteil an, nicht jedoch auf den Rest des Hauptteils, der diesen Grenzwert überschreitet (die tatsächliche Anzahl der Zeichen hängt von der Codierung ab – diese kann 1 bis 4 Byte pro Zeichen betragen).|Wendet den regulären Ausdruck auf die ersten 64.000 Zeichen (inkl. HTML-Tag-Zeichen) an, nicht jedoch auf den Rest des Hauptteils, der diesen Grenzwert überschreitet.|Aktiviert das Add-in nur, wenn der Hauptteil < 500.000 Zeichen groß ist.|

    In Tabelle 3 werden die Grenzwerte und die Unterschiede hinsichtlich der Treffer erläutert, die von jedemOutlook-Host nach der Überprüfung eines regulären Ausdrucks zurückgegeben werden. Die Unterstützung ist nicht abhängig von bestimmten Arten von Geräten, hängt jedoch möglicherweise von der Art des Hauptteils des Elements ab, wenn der reguläre Ausdruck auf den Hauptteil des Elements angewendet wird.
    

    **Tabelle 3. Grenzwerte für die zurückgegebenen Treffer**


||**Outlook-Rich-Client**|**Outlook Web App oder OWA für mobile Geräte**|
|:-----|:-----|:-----|
|Reihenfolge der zurückgegebenen Übereinstimmungen|Gehen Sie davon aus, dass  **getRegExMatches** für den gleichen regulären Ausdruck, der auf das gleiche Element angewendet wird, in einem Outlook-Rich-Client andere Übereinstimmungen zurückgibt als in Outlook Web App oder OWA für mobile Geräte.|Gehen Sie davon aus, dass  **getRegExMatches** Übereinstimmungen in einem Outlook-Rich-Client in anderer Reihenfolge zurückgibt als in Outlook Web App oder OWA für mobile Geräte.|
|Nur-Text-Hauptteil des Elements|**getRegExMatches** gibt alle Treffer bis zu einer Größe von 1.536 Zeichen (1,5 KB) zurück (maximal 50 Treffer).
(../../images   **getRegExMatches** gibt Treffer im zurückgegebenen Array nicht in einer bestimmten Reihenfolge zurück. Im Allgemeinen können Sie davon ausgehen, dass sich bei Anwendung des gleichen regulären Ausdrucks auf das gleiche Element die Reihenfolge der Treffer, die für einen Outlook-Rich-Client zurückgegeben werden, von der Reihenfolge der Treffer unterscheidet, die für Outlook Web App und OWA für mobile Geräte zurückgegeben werden.

|**getRegExMatches** gibt alle Treffer bis zu einer Größe von 3.072 Zeichen (3 KB) zurück (maximal 50 Treffer).|
|Hauptteil des HTML-Elements|**getRegExMatches** gibt alle Treffer bis zu einer Größe von 3.072 Zeichen (3 KB) zurück (maximal 50 Treffer).
(../../images   **getRegExMatches** gibt Treffer im zurückgegebenen Array nicht in einer bestimmten Reihenfolge zurück. Im Allgemeinen können Sie davon ausgehen, dass sich bei Anwendung des gleichen regulären Ausdrucks auf das gleiche Element die Reihenfolge der Treffer, die für einen Outlook-Rich-Client zurückgegeben werden, von der Reihenfolge der Treffer unterscheidet, die für Outlook Web App und OWA für mobile Geräte zurückgegeben werden.

|**getRegExMatches** gibt alle Treffer bis zu einer Größe von 3.072 Zeichen (3 KB) zurück (maximal 50 Treffer).|

## Grenzwerte für JavaScript-API


Außer den o. g. Richtlinien für Aktivierungsregeln erzwingt jeder einzelne Outlook-Host bestimmte Grenzwerte hinsichtlich des JavaScript-Objektmodells, die in Tabelle 4 beschrieben werden.


**Tabelle 4. Grenzwerte für das Abrufen oder Festlegen bestimmter Daten über die JavaScript-API für Office**


|**Funktion**|**Grenzwert**|**Zugehörige API**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|Benutzerdefinierte Eigenschaften|2.500 Zeichen|[CustomProperties](http://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)-Objekt [item.loadCustomPropertiesAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode|Grenzwert für alle benutzerdefinierten Eigenschaften für ein Termin- oder Nachricht-Objekt. Alle Outlook-Hosts geben einen Fehler zurück, wenn die Gesamtgröße aller benutzerdefinierten Eigenschaften eines Outlook-Add-Ins diesen Grenzwert überschreitet.|
|Roamingeinstellungen|Zeichenanzahl: 32 KB|[RoamingSettings](http://dev.outlook.com/reference/add-ins/RoamingSettings.html%28Office.15%29.md)-Objekt [context.roamingSettings](http://dev.outlook.com/reference/add-ins/Office.context.html%28Office.15%29.md)-Eigenschaft|Grenzwert für alle Roamingeinstellungen des Outlook-Add-Ins. Alle Outlook-Hosts geben einen Fehler zurück, wenn die Einstellungen diesen Grenzwert überschreiten.|
|Extrahieren bekannter Entitäten|Zeichenanzahl: 2.000|[item.getEntities](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode [item.getEntitiesByType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode [item.getFilteredEntitiesByName](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode|Grenzwert für Exchange Server, um bekannte Entitäten im Hauptteil zu extrahieren. Exchange Server ignoriert über diesen Grenzwert hinausgehende Entitäten. Beachten Sie, dass dieser Grenzwert unabhängig davon ist, ob das Outlook-Add-In eine  **ItemHasKnownEntity**-Regel verwendet.|
|Exchange-Webdienste|Zeichenanzahl: 1 MB|[mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode|Grenzwert für eine Anforderung oder eine Antwort an einen  **Mailbox.makeEwsRequestAsync**-Aufruf.|
|Empfänger|100 Empfänger|[item.requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft [item.optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft [item.resources](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft [item.to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft [item.cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Eigenschaft [Recipients.addAsync ](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Methode [Recipient.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Methode [Recipient.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Methode|Grenzwert für die Empfänger, die in jeder Eigenschaft angegeben sind.|
|Anzeigename|255 Zeichen|[EmailAddressDetails.displayName ](http://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Eigenschaft [Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Objekt **item.requiredAttendees**-Eigenschaft **item.optionalAttendees**-Eigenschaft **item.resources**-Eigenschaft **item.to**-Eigenschaft **item.cc**-Eigenschaft|Grenzwert für die Länge eines Anzeigenamens in einem Termin oder einer Nachricht.|
|Festlegen des Betreffs|255 Zeichen|[mailbox.displayNewAppointmentForm](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode [Subject.setAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)-Methode|Grenzwert für den Betreff im neuen Terminformular oder für das Festlegen des Betreffs eines Termins oder einer Nachricht.|
|Festlegen des Ortes|255 Zeichen|[Location.setAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)-Methode|Grenzwert für das Festlegen des Ortes eines Termins oder einer Besprechungsanfrage.|
|Hauptteil in einem neuen Terminformular|Zeichenanzahl: 32 KB|Methode  **Mailbox.displayNewAppointmentForm**|Grenzwert für den Hauptteil in einem neuen Terminformular.|
|Anzeigen des Hauptteils eines vorhandenen Objekts|Zeichenanzahl: 32 KB|[mailbox.displayAppointmentForm](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode [mailbox.displayMessageForm](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode|Für Outlook Web App und OWA für mobile Geräte: Grenzwert für den Hauptteil in einem vorhandenen Termin- oder Nachrichtenformular.|
|Festlegen des Hauptteils|Zeichenanzahl: 1 MB|[Body.prependAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)-Methode [Body.setAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)[Body.setSelectedDataAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)-Methode|Grenzwert für das Festlegen des Hauptteils eines Termin- oder Nachrichtenobjekts.|
|Anzahl der Anlagen|499 Dateien in Outlook Web App und OWA für mobile Geräte|[item.addFileAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode|Grenzwert für die Anzahl von Dateien, die an ein Element beim Senden angehängt werden können. Für Outlook Web App und OWA für mobile Geräte liegt dieser Grenzwert im Allgemeinen bei 499 Dateien über die Benutzeroberfläche und  **addFileAttachmentAsync**. Ein Outlook-Rich-Client besitzt keinen spezifischen Grenzwert für die Anzahl der Dateianlagen. Alle Outlook-Hosts überwachen jedoch den Grenzwert für die Anlagengröße, die für den Exchange Server des Benutzers konfiguriert ist. Siehe "Anlagengröße" in der nächsten Zeile.|
|Anlagengröße|Abhängig von Exchange Server|**item.addFileAttachmentAsync**-Methode|Es besteht ein Grenzwert für die Größe aller Anlagen eines Objekts, den ein Administrator im Exchange Server des Postfach des Benutzers konfigurieren kann.Bei einem Outlook-Rich-Client gilt dieser Grenzwert für die Anzahl von Anlagen für ein Element. Bei Outlook Web App und OWA für mobile Geräte gilt jeweils der geringere der beiden Grenzwerte – die Anzahl von Anlagen und die Größe aller Anlagen – für die Beschränkung der tatsächlich für ein Element vorhandenen Anlagen.|
|Dateiname der Anlage|255 Zeichen|**item.addFileAttachmentAsync**-Methode|Grenzwert für die Länge des Dateinamens einer einem Objekt hinzuzufügenden Anlage.|
|Anlagen-URI|2048 Zeichen|**item.addFileAttachmentAsync**-Methode|Grenzwert für den URI des Dateienamens, der einem Objekt als Anlage hinzugefügt werden soll.|
|Anlagen-ID|100 Zeichen|[item.addItemAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode [item.removeAttachmentAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode|Grenzwert für die ID der Anlage, die einem Objekt hinzugefügt oder von ihm entfernt werden soll.|
|Asynchrone Aufrufe|3 Aufrufe|**item.addFileAttachmentAsync**-Methode **item.addItemAttachmentAsync**-Methode **item.removeAttachmentAsync**-Methode [Body.prependAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)-MethodeMethode  **Body.prependAsync**Methode  **Body.setSelectedDataAsync**[CustomProperties.saveAsync](http://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)-Methode [item.LoadCustomPropertiesAysnc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)-Methode [Location.getAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)-MethodeMethode  **Location.setAsync**[mailbox.getCallbackTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode [mailbox.getUserIdentityTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-MethodeMethode  **Recipients.addAsync**[Recipient.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)-Methode **Recipients.setAsync**-Methode [RoamingSettings.saveAsync](http://dev.outlook.com/reference/add-ins/RoamingSettings.html%28Office.15%29.md)-Methode [Subject.setAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)-MethodeMethode  **Subject.setAsync**[Time.getAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)-Methode [Time.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)-Methode|Bei Outlook Web App oder OWA für mobile Geräte: Grenzwert für die Anzahl gleichzeitiger asynchroner Aufrufe, da Browser nur eine begrenzte Anzahl asynchroner Aufrufe für Server zulassen. |

## Zusätzliche Ressourcen



- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](d6eea4c4-bb21-4f24-bcba-1eccbb4e12dd.md)
    
- [Datenschutz, Berechtigungen und Sicherheit für Outlook-Add-Ins](44208fc4-05d4-42d8-ab20-faa89624de1c.md)
    
