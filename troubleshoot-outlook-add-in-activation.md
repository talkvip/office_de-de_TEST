
# Problembehandlung für die Aktivierung von Outlook-Add-Ins


Die Aktivierung des Outlook-Add-Ins erfolgt nach Kontext und basiert auf den Aktivierungsregeln im Add-In-Manifest. Wenn die Bedingungen für das momentan ausgewählte Element die Aktivierungsregeln für das Add-In erfüllen, wird die Hostanwendung aktiviert und zeigt die Add-In-Schaltfläche in der Benutzeroberfläche von Outlook an (Add-In-Auswahlbereich bei Verfassen-Add-Ins, Add-In-Leiste bei Lese-Add-Ins). Wird das Add-In jedoch nicht erwartungsgemäß aktiviert, sollten Sie in den folgenden Bereichen nach der Ursache suchen.

<a name="troubleshootingmailapps"></a>
## Befindet sich das Benutzerpostfach unter einer Version von Exchange Server, die mindestens Exchange 2013 lautet?


Stellen Sie zuerst sicher, dass sich das E-Mail-Konto des Benutzers, mit dem Sie den Test durchführen, mindestens unter Version Exchange 2013 befindet. Bei Verwendung bestimmter Funktionen, die nach Exchange 2013 veröffentlicht wurden, vergewissern Sie sich, dass sich das Benutzerkonto auf der richtigen Version von Exchange befindet.

Sie können die Version von Exchange 2013 mit einer der folgenden Vorgehensweisen überprüfen:


- Fragen Sie Ihren Exchange Server-Administrator.
    
- Wenn Sie das Add-In unter Outlook Web App oder OWA für mobile Geräte testen, sollten Sie in einem Skriptdebugger (z. B. im JScript-Debugger von Internet Explorer) nach dem  **src**-Attribut des  **script**-Tags suchen. Damit wird der Speicherort angegeben, von dem Skripts geladen werden. Der Pfad sollte die Unterzeichenfolge  **owa/15.0.516.x/owa2/…** enthalten, wobei **15.0.516.x** für die Exchange Server-Version steht, z. B. **15.0.516.2**.
    
- Alternativ dazu können Sie die [Office.context.mailbox.diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md)-Eigenschaft, um die Version zu überprüfen. Unter Outlook Web App und OWA für mobile Geräte gibt diese Eigenschaft die Version von Exchange Server zurück.
    
- Wenn Sie das Add-In in Outlook testen können, können Sie die folgende einfache Debugtechnik nutzen, bei der das Outlook-Objektmodell und Visual Basic Editor verwendet werden:
    
      1. First, verify that macros are enabled for Outlook. Choose  **File**,  **Options**,  **Trust Center**,  **Trust Center Settings**,  **Macro Settings**. Ensure that  **Notifications for all macros** is selected in the Trust Center. You should have also selected **Enable Macros** during Outlook startup.
    
      2. On the  **Developer** tab of the ribbon, choose **Visual Basic**.
    
     >
  **Note**  Not seeing the  **Developer** tab? See [How to: Show the Developer Tab on the Ribbon](http://msdn.microsoft.com/en-us/library/ce7cb641-44f2-4a40-867e-a7d88f8e98a9%28Office.15%29.aspx) to turn it on. 3. Wählen Sie im Visual Basic-Editor die Optionen  **Ansicht** und **Direktfenster**.
    
      4. Type the following in the Immediate window to display the version of the Exchange Server. The major version of the returned value must be equal to or greater than 15.
    
        - If there is only one Exchange account in the user's profile:
        
            
            ?Session.ExchangeMailboxServerVersion
            
        
        - If there are multiple Exchange accounts in the same user profile:
        
            
            ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
         
        
        - _emailAddress_ represents a string that contains the user's primary STMP address. For example, if the user's primary SMTP address is randy@contoso.com, type the following:
        
            
            ?Session.Accounts.Item("randy@contoso.com").ExchangeMailboxServerVersion
        


## Ist das Add-In deaktiviert?


Ein beliebiger der Outlook-Rich-Clients kann ein Add-In aus Leistungsgründen deaktivieren, z. B. bei Überschreitung der Schwellenwerte für die Nutzung von CPU-Core oder Arbeitsspeicher, der Absturztoleranz und des Zeitraums der Verarbeitung aller regulären Ausdrücke für ein Mail-Add-in. In diesem Fall zeigt der Outlook-Rich-Client eine Benachrichtigung an, dass das Add-In deaktiviert wird. 


 >**Note**  Only Outlook rich clients monitor resource usage, but disabling an add-in in an Outlook rich client also disables the add-in in Outlook Web App and OWA for Devices.

Verwenden Sie eines der folgenden Verfahren, um zu überprüfen, ob ein Add-In deaktiviert ist: 


- Melden Sie sich in Outlook Web App direkt am E-Mail-Konto an, wählen Sie das Symbol „Einstellungen" aus, und klicken Sie auf  **Add-ins verwalten**, um auf die Exchange-Verwaltungskonsole (Exchange Admin Center) zuzugreifen. Darin können Sie überprüfen, ob das Add-in deaktiviert ist.
    
- Greifen Sie in Outlook auf die Backstage-Ansicht zu, und wählen Sie  **Add-Ins verwalten** aus. Melden Sie sich beim Exchange Admin Center an, um zu prüfen, ob das Add-In aktiviert ist.
    
- Wählen Sie in Outlook für Mac in der Add-in-Leiste  **Add-ins verwalten** aus. Melden Sie sich beim Exchange Admin Center an, um zu prüfen, ob das Add-in aktiviert ist.
    

## Unterstützen die getesteten Elemente Outlook-Add-Ins? Wird das ausgewählte Objekt von einer Exchange Server-Version bereitgestellt, die mindestens Exchange 2013 lautet?


Wenn es sich bei dem Outlook-Add-In um ein Lese-Add-In handelt und dieses aktiviert werden soll, sobald ein Benutzer eine Nachricht (z. B. E-Mail-Nachrichten, Besprechungsanfragen, Antworten und Absagen) oder einen Termin anzeigt, kann es zu Ausnahmen kommen, obwohl diese Elemente Add-Ins normalerweise unterstützen. Ausnahmen treten auf, wenn für das ausgewählte Objekt Folgendes zutrifft.


- Per Verwaltung von Informationsrechten (IRM) geschützt.
    
- S/MIME-Format oder aus Schutzgründen auf andere Art und Weise verschlüsselt.
    
- Entwurf (kein Absender zugewiesen) oder im Outlook-Ordner "Entwürfe" enthalten.
    
- Im Ordner "Junk-E-Mail" enthalten.
    
- Übermittlungsbericht oder -benachrichtigung mit der Nachrichtenklasse "IPM.Report.*", darunter Übermittlungs- und Unzustellbarkeitsberichte sowie Benachrichtigungen über "Gelesen", "Ungelesen" und "Verzögerung".
    
- MSG-Dateien, die als Anlage für eine andere Nachricht fungieren oder im Dateisystem geöffnet werden.
    
Da Termine immer im Rich-Text-Format gespeichert werden, aktiviert eine [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regel, mit der der  **PropertyName**-Wert  **BodyAsHTML** angegeben wird, kein Add-In für einen Termin oder eine Nachricht, der bzw. die im Nur-Text- oder Rich-Text-Format gespeichert ist..

Auch wenn ein Mail-Element nicht einen der obigen Typen aufweist und das Element nicht mindestens von der Exchange Server-Version Exchange 2013 bereitgestellt wird, werden bekannte Entitäten und Eigenschaften wie die SMTP-Adresse des Absenders für das Element nicht identifiziert. Alle Aktivierungsregeln, die auf diesen Entitäten oder Eigenschaften basieren, wären nicht erfüllt, und das Add-In wird somit nicht aktiviert.

Wenn es sich bei dem Add-In um ein Verfassen-Add-In handelt und es beim benutzerseitigen Verfassen einer Nachricht oder einer Besprechungsanfrage aktiviert werden soll, vergewissern Sie sich, dass das Objekt nicht durch IRM geschützt wird.


## Ist das Add-in-Manifest ordnungsgemäß installiert, und verfügt Outlook über eine Kopie im Cache?


Dieses Szenario gilt nur für Outlook für Windows. Normalerweise kopiert der Exchange Server das Add-in-Manifest beim Installieren eines Outlook-Add-Ins für ein Postfach vom angegebenen Speicherort in das jeweilige Postfach für Exchange Server. Bei jedem Start von Outlook liest Outlook alle Manifeste, die für das Postfach installiert sind, in einen temporären Cache am folgenden Speicherort ein: 

%LocalAppData%\Microsoft\Office\15.0\WEF 

Für den Benutzer Jan kann sich der Cache z. B. unter "C:\Users\jan\AppData\Local\Microsoft\Office\15.0\WEF" befinden.

Wenn ein Add-In für Elemente nicht aktiviert wird, wurde das Manifest auf dem Exchange Server möglicherweise nicht richtig installiert, oder Outlook hat das Manifest beim Starten nicht richtig gelesen. Stellen Sie mithilfe der Exchange-Verwaltungskonsole sicher, dass das Add-In installiert und für Ihr Postfach aktiviert ist. Starten Sie den Exchange Server neu, falls dies erforderlich ist.

In Abbildung 1 ist eine Übersicht der Schritte dargestellt, mit denen Sie prüfen können, ob Outlook über eine gültige Version des Manifests verfügt. 


**Abbildung 1: Flussdiagramm der Schritte zum Überprüfen, ob das Manifest von Outlook richtig zwischengespeichert wurde**

![Flussdiagramm für Manifestüberprüfung](../../images/off15appsdk_TroubleshootManifest.png)Das folgende Verfahren enthält eine ausführliche Beschreibung.



1. Wenn Sie das Manifest bei geöffneter Outlook-Anwendung geändert haben und Napa, Visual Studio 2012 oder eine spätere Version von Visual Studio nicht zum Entwickeln desAdd-Ins verwenden, sollten Sie das Add-In deinstallieren und mithilfe der Exchange-Verwaltungskonsole neu installieren. 
    
2. Starten Sie Outlook neu, und testen Sie, ob das Add-in von Outlook jetzt aktiviert wird.
    
3. Falls das Add-in von Outlook nicht aktiviert wird, sollten Sie überprüfen, ob Outlook über eine ordnungsgemäß zwischengespeicherte Kopie des Manifests für das Add-in verfügt. Sehen Sie unter folgendem Pfad nach:
    
    %LocalAppData%\Microsoft\Office\15.0\WEF
    
    Sie finden das Manifest im folgenden Unterordner:
```
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
```
    
     >**Note**  The following is an example of a path to a manifest installed for a mailbox for the user John:
    
    C:\Users\john\appdata\Local\Microsoft\Office\15.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    
    Verify whether the manifest of the add-in you're testing is among the cached manifests.
    
4. Wenn das Manifest im Cache enthalten ist, können Sie den Rest dieses Abschnitts überspringen und sich den anderen möglichen Ursachen widmen, die nach diesem Abschnitt beschrieben werden.
    
5. Wenn das Manifest nicht im Cache enthalten ist, sollten Sie prüfen, ob Outlook das Manifest wirklich aus dem Exchange Server ausgelesen hat. Verwenden Sie dazu die Windows-Ereignisanzeige:
    
      1. Wählen Sie unter  **Windows-Protokolle** die Option **Anwendung**.
    
      2. Suchen Sie nach einem kürzlich eingetretenen Ereignis mit der Ereignis-ID 63. Diese ID zeigt an, dass Outlook ein Manifest aus einem Exchange Server heruntergeladen hat.
    
      3. Wenn Outlook ein Manifest erfolgreich gelesen hat, sollte das protokollierte Ereignis über eine Beschreibung der folgenden Art verfügen:
    
         **Die Exchange-Webdienstanforderung „GetAppManifests“ war erfolgreich.**
    
        Sie können den Rest dieses Abschnitts überspringen und sich den anderen möglichen Ursachen widmen, die nach diesem Abschnitt beschrieben werden.
    

    Informationen zum Öffnen der Ereignisanzeige in Windows 7 finden Sie unter [Öffnen der Ereignisanzeige](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer).
    
6. Falls Sie kein erfolgreiches Ereignis finden können, schließen Sie Outlook, und löschen Sie alle Manifeste unter dem folgenden Pfad:
```
    %LocalAppData%\Microsoft\Office\15.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
```
    Start Outlook and test whether Outlook now activates the add-in.
    
7. Wenn das Add-in von Outlook nicht aktiviert wird, gehen Sie zurück zu Schritt 3, um erneut zu überprüfen, ob Outlook das Manifest richtig gelesen hat.
    

## Verwenden Sie die passenden Aktivierungsregeln?


Ab Version 1.1 des Office-Add-Ins-Manifestschemas können Sie Add-Ins erstellen, die aktiviert werden, sobald sich der Benutzer in einem Formular zum Verfassen (Verfassen-Add-Ins) oder in einem Leseformular (Lese-Add-Ins) befindet. Vergewissern Sie sich, dass Sie für alle Typen oder Formulare, in denen Ihr Add-In jeweils aktiviert werden soll, die passenden Aktivierungsregeln festlegen. Das Aktivieren von Verfassen-Add-Ins ist beispielsweise nur mit [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)-Regeln möglich, deren  **FormType**-Attribut auf  **Edit** oder **ReadOrEdit** gesetzt ist. Das Verwenden anderer Regeltypen wie [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)- und [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)-Regeln ist für Verfassen-Add-ins nicht möglich. Weitere Informationen finden Sie unter [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md).


## Ist der reguläre Ausdruck richtig angegeben (falls verwendet)?


Da reguläre Ausdrücke in Aktivierungsregeln Teil der XML-Manifestdatei eines Lese-Add-ins sind, sollten Sie, falls in einem regulären Ausdruck entsprechende Zeichen verwendet werden, die dafür geltende Escapesequenz befolgen, die von XML-Prozessoren unterstützt wird. In Tabelle 1 sind diese Sonderzeichen aufgeführt. 


**Tabelle 1: Escapesequenzen für reguläre Ausdrücke**


|**Zeichen**|**Beschreibung**|**Zu verwendende Escapesequenz**|
|:-----|:-----|:-----|
|"|Doppeltes Anführungszeichen|&amp;quot;|
|&amp;|Kaufmännisches Und-Zeichen|&amp;amp;|
|'|Apostroph|&amp;apos;|
|<|Kleiner als-Zeichen|&amp;lt;|
|>|Größer als-Zeichen|&amp;gt;|

## Wenn Sie einen regulären Ausdruck verwenden, wird das Lese-Add-in dann in Outlook Web App oder OWA für mobile Geräte aktiviert, jedoch in keinem der Outlook-Rich-Clients?


Outlook-Rich-Clients verwenden ein Modul für reguläre Ausdrücke, das sich von demjenigen unterscheidet, das von Outlook Web App und OWA für mobile Geräte verwendet wird. Outlook-Rich-Clients verwenden das C++-Modul für reguläre Ausdrücke, das als Teil der standardmäßigen Vorlagenbibliothek von Visual Studio bereitgestellt wird. Dieses Modul erfüllt die ECMAScript 5-Normen. Outlook Web App und OWA für mobile Geräte verwenden eine Auswertung für reguläre Ausdrücke, die Teil von JavaScript ist, über den Browser bereitgestellt wird und eine Obermenge von ECMAScript 5 darstellt. 

In den meisten Fällen finden diese Hostanwendungen für einen regulären Ausdruck zwar die gleichen Übereinstimmungen in einer Aktivierungsregel, aber es gibt auch Ausnahmen. Wenn der reguläre Ausdruck beispielsweise eine benutzerdefinierte Zeichenklasse enthält, die auf vordefinierten Zeichenklassen basiert, werden für einen Outlook-Rich-Client möglicherweise andere Ergebnisse als für Outlook Web App und OWA für mobile Geräte zurückgeben. Für Zeichenklassen, die Kurzschrift-Zeichenklassen  `[\d\w]` enthalten, werden z. B. unterschiedliche Ergebnisse zurückgegeben. In diesem Fall sollten Sie stattdessen `(\d|\w)` verwenden, um unterschiedliche Ergebnisse für verschiedene Hosts zu vermeiden.

Testen Sie Ihre regulären Ausdrücke sorgfältig. Falls unterschiedliche Ergebnisse zurückgegeben werden, sollten Sie den regulären Ausdruck so umschreiben, dass er mit beiden Modulen kompatibel ist. Um die Auswertungsergebnisse für einen Outlook-Rich-Client zu überprüfen, schreiben Sie ein kurzes C++-Programm, mit dem der reguläre Ausdruck auf einen Auszug des Texts angewendet wird, für den Sie eine Übereinstimmung erzielen möchten. Bei Ausführung unter Visual Studio verwendet das C++-Testprogramm die standardmäßige Vorlagenbibliothek und simuliert das Verhalten des Outlook-Rich-Clients für denselben regulären Ausdruck. Verwenden Sie Ihr bevorzugtes JavaScript-Testprogramm für reguläre Ausdrücke, um die Auswertungsergebnisse für Outlook Web App oder OWA für mobile Gerätezu überprüfen.


## Haben Sie bei Verwendung einer ItemIs-, ItemHasAttachment- oder ItemHasRegularExpressionMatch-Regel die zugehörige Elementeigenschaft überprüft?


Überprüfen Sie bei Verwendung einer  **ItemHasRegularExpressionMatch**-Aktivierungsregel, ob der Wert des  **PropertyName**-Attributs für das ausgewählte Element wie erwartet ist. Es folgen einige Tipps zum Debuggen der entsprechenden Eigenschaften:


- Wenn es sich beim ausgewählten Element um eine Nachricht handelt und Sie  **BodyAsHTML** im **PropertyName**-Attribut angeben, öffnen Sie die Nachricht, und wählen Sie  **Quelle anzeigen**, um den Nachrichtentext in der HTML-Darstellung des Elements zu überprüfen.
    
- Wenn es sich beim ausgewählten Element um einen Termin handelt oder wenn in der Aktivierungsregel  **BodyAsPlaintext** unter **PropertyName** angegeben ist, können Sie das Outlook-Objektmodell und den Visual Basic Editor in Outlook für Windows verwenden:
    
      1. Ensure that macros are enabled and the  **Developer** tab is displayed in the ribbon for Outlook. If you're not sure how to do this, see Steps 1 and 2 under [Is the user mailbox on a version of Exchange Server that is at least Exchange 2013?](#is-the-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013)
    
      2. In the Visual Basic Editor, choose  **View**,  **Immediate Window**.
    
      3. Type the following to display various properties depending on the scenario. 
    
      - The HTML body of the message or appointment item selected in the Outlook explorer:
    
            
              ?ActiveExplorer.Selection.Item(1).HTMLBody
        


     - "Nur-Text" des im Outlook Explorer ausgewählten Nachrichten- oder Terminelements:
    
            
              ?ActiveExplorer.Selection.Item(1).Body
            


      - HTML-Text des im aktuellen Outlook-Inspector geöffneten Nachrichten- oder Terminelements:
    
            
              ?ActiveInspector.CurrentItem.HTMLBody
        
      - "Nur-Text" des im aktuellen Outlook-Inspector geöffneten Nachrichten- oder Terminelements:
    
            
              ?ActiveInspector.CurrentItem.Body
            

Wenn in der  **ItemHasRegularExpressionMatch**-Aktivierungsregel  **Subject** oder **SenderSMTPAddress** angegeben ist oder wenn Sie eine **ItemIs**- oder  **ItemHasAttachment**-Regel verwenden und mit MAPI vertraut sind bzw. MAPI nutzen möchten, können Sie [MFCMAPI](http://mfcmapi.codeplex.com/) einsetzen. Darüber können Sie den Wert in Tabelle 2 überprüfen, auf dem Ihre Regel beruht.


**Tabelle 2: Aktivierungsregeln und dazugehörige MAPI-Eigenschaften**


|**Regeltyp**|**Zu überprüfende MAPI-Eigenschaft**|
|:-----|:-----|
|**ItemHasRegularExpressionMatch**-Regel mit  **Subject**|[PidTagSubject](http://msdn.microsoft.com/en-us/library/aa7ba4d9-c5e0-4ce7-a34e-65f675223bc9%28Office.15%29.aspx)|
|**ItemHasRegularExpressionMatch**-Regel mit  **SenderSMTPAddress**|
  [PidTagSenderSmtpAddress](http://msdn.microsoft.com/en-us/library/321cde5a-05db-498b-a9b8-cb54c8a14e34%28Office.15%29.aspx) und [PidTagSentRepresentingSmtpAddress](http://msdn.microsoft.com/en-us/library/5ed122a2-0967-4de3-a2ee-69f81ae77b16%28Office.15%29.aspx)|
|**ItemIs**|[PidTagMessageClass](http://msdn.microsoft.com/en-us/library/1e704023-1992-4b43-857e-0a7da7bc8e87%28Office.15%29.aspx)|
|**ItemHasAttachment**|[PidTagHasAttachments](http://msdn.microsoft.com/en-us/library/fd236d74-2868-46a8-bb3d-17f8365931b6%28Office.15%29.aspx)|
Nach dem Überprüfen des Eigenschaftswerts können Sie ein Tool zur Auswertung regulärer Ausdrücke verwenden, um zu testen, ob der reguläre Ausdruck eine Übereinstimmung für diesen Wert ergibt.


## Wendet die Hostanwendung alle regulären Ausdrücke wie erwartet auf den Teil des Elementtextkörpers an?


Dieser Abschnitt gilt für alle Aktivierungsregeln, in denen reguläre Ausdrücke verwendet werden. Dies gilt besonders für die Ausdrücke, die auf den Elementtext angewendet werden und relativ groß sein können, sodass die Auswertung mit der Rückgabe von Übereinstimmungen möglicherweise länger dauert. Beachten Sie Folgendes: Auch wenn die Elementeigenschaft, von der eine Aktivierungsregel abhängt, den erwarteten Wert aufweist, kann die Hostanwendung möglicherweise nicht alle regulären Ausdrücke für den gesamten Wert der Elementeigenschaft auswerten. Um eine angemessene Leistung zu erzielen und den übermäßigen Ressourceneinsatz eines Lese-Add-ins zu verhindern, gelten für Outlook, Outlook Web App und OWA für mobile Geräte in Bezug auf die Verarbeitung regulärer Ausdrücke in Aktivierungsregeln zur Laufzeit die folgenden Grenzwerte:


- Größe des auszuwertenden Elementtexts: Es gelten Grenzwerte für den Teil eines Elementtexts, für den die Hostanwendung einen regulären Ausdruck auswertet. Diese Grenzwerte richten sich jeweils nach der Hostanwendung, dem Formularfaktor und dem Format des Elementtexts. Ausführliche Informationen finden Sie in Tabelle 2 unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
- Anzahl der Übereinstimmungen für reguläre Ausdrücke: Die Outlook-Rich-Clients, Outlook Web App und OWA für mobile Geräte geben jeweils maximal 50 Übereinstimmungen für reguläre Ausdrücke zurück. Diese Übereinstimmungen sind eindeutig, und doppelte Übereinstimmungen werden bei diesem Grenzwert nicht einberechnet. Für die zurückgegebenen Übereinstimmungen gilt keine bestimmte Reihenfolge, und es ist nicht sichergestellt, dass die Reihenfolge in einem Outlook-Rich-Client und in Outlook Web App und OWA für mobile Geräte damit identisch ist. Falls Sie für die regulären Ausdrücke in Ihren Aktivierungsregeln viele Übereinstimmungen erwarten und eine Übereinstimmung fehlt, können Sie diesen Grenzwert erhöhen.
    
- Länge der Übereinstimmung für einen regulären Ausdruck: Für die Rückgabe von Übereinstimmungen für einen regulären Ausdruck gilt eine Längenbeschränkung. In der Hostanwendung werden keine Übereinstimmungen einbezogen, die oberhalb des Grenzwerts liegen, und es wird keine Warnmeldung angezeigt. Sie können Ihren regulären Ausdruck mithilfe anderer Tools zur Auswertung regulärer Ausdrücke oder eines eigenständigen C++-Testprogramms ausführen, um zu überprüfen, ob einige Übereinstimmungen den Grenzwert überschreiten. In Tabelle 3 sind die Grenzwerte zusammengefasst. Weitere Informationen finden Sie in Tabelle 3 unter [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
    **Tabelle 3. Längenbeschränkungen für eine Übereinstimmung für einen regulären Ausdruck**


|**Grenzwert für eine Übereinstimmung für einen regulären Ausdruck**|**Outlook-Rich-Clients**|**Outlook Web App oder OWA für mobile Geräte**|
|:-----|:-----|:-----|
|Elementtext ist Nur-Text|1,5 KB|3 KB|
|Elementtext ist HTML|3 KB|3 KB|
- Zeitaufwand für Auswertung aller regulären Ausdrücke eines Lese-Add-ins für einen Outlook-Rich-Client: Standardmäßig muss Outlook die Auswertung aller regulären Ausdrücke innerhalb der Aktivierungsregeln für jedes Mail-Add-in innerhalb einer Sekunde abschließen. Andernfalls versucht Outlook bis zu dreimal, den Vorgang zu wiederholen, und deaktiviert das Add-in, falls die Auswertung von Outlook nicht abgeschlossen werden kann. In Outlook wird in der Benachrichtigungsleiste eine Meldung angezeigt, dass das Add-in deaktiviert wurde. Wie viel Zeit für den regulären Ausdruck verfügbar ist, können Sie einstellen, indem Sie eine Gruppenrichtlinie oder einen Registrierungsschlüssel festlegen. 
    
     >**Hinweis**  Beachten Sie, dass, wenn der Outlook-Rich-Client ein Lese-Add-In deaktiviert, das Lese-Add-In nicht für die Verwendung für dasselbe Postfach im Outlook-Rich-Client, in der Outlook Web App und OWA für Geräte verfügbar ist.

## Weitere Ressourcen



- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md)
    
- [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md)
    
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Einschränkungen für die Aktivierung und die JavaScript-API für Outlook-Add-Ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Öffnen der Ereignisanzeige](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer)
    
- [ItemHasAttachment complexType](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)
    
- [ItemHasRegularExpressionMatch complexType](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)
    
- [ItemIs complexType](http://msdn.microsoft.com/en-us/library/926249ab-2d2f-39f5-1d73-fab1c989966f%28Office.15%29.aspx)
    
- [MailApp complexType](http://msdn.microsoft.com/en-us/library/696b9fcf-cd10-3f20-4d49-86d3690c887a%28Office.15%29.aspx)
    
