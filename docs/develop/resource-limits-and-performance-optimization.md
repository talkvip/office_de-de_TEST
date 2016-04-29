
# Ressourcengrenzen und Leistungsoptimierung für Office Add-Ins
Wenden Sie Ressourcen-Nutzungsbegrenzungen auf Ihre Office-Add-Ins an und optimieren Sie die Leistung Ihres Add-Ins. 

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Um für Ihre Benutzer eine höchstmögliche Benutzerfreundlichkeit zu gewährleisten, müssen Sie sicherstellen, dass Ihre Office-Add-In im Rahmen bestimmter Grenzwerte für die Auslastung der CPU-Kerne und des Arbeitsspeichers, für die Zuverlässigkeit und im Fall von Outlook-Add-Ins für die Antwortzeit für die Überprüfung regulärer Ausdrücke funktioniert. Diese Grenzwerte für die Nutzung von Laufzeitressourcen gelten für die Add-Ins auf Office-Clients unter Windows und OS X, jedoch nicht für Office Online, Outlook Web App oder OWA für mobile Geräte. Sie können außerdem die Leistung Ihrer Add-Ins auf Desktop- und mobilen Geräten verbessern, indem Sie die Verwendung der Ressourcen in Ihrem Add-In-Design und der Implementierung optimieren.

## Grenzwerte für die Ressourcennutzung für Add-Ins


Durch Grenzwerte für die Nutzung von Laufzeitressourcenwenden alle Typen von Office-Add-Ins an. Diese Grenzwerte gewährleisten ausreichend Leistung für die Benutzer und tragen außerdem dazu bei, das Risikopotenzial von Denial-of-Service-Angriffen zu verringern. Testen Sie Ihre Office-Add-In auf dem Ziel-Host mit vielen verschiedenen Daten, messen Sie die Leistung im Hinblick auf die folgenden Runtime-Grenzwerte und Nutzungsbegrenzungen.


-  **CPU-Kern-Auslastung** – Ein Grenzwert für die Auslastung eines einzelnen CPU-Kerns in Höhe von 90 % (drei Mal in 5-Sekunden-Standardintervallen beobachtet).
    
    Das Standardintervall, in dem ein Host-Rich-Client die CPU-Kern-Auslastung überprüft, beträgt 5 Sekunden. Stellt der Host-Client fest, dass die CPU-Kern-Auslastung durch ein Add-In sich über diesem Grenzwert befindet, wird eine Meldung angezeigt, in der der Benutzer gefragt wird, ob er das Add-In weiterhin ausführen möchte. Führt der Benutzer das Add-In weiterhin aus, wird er vom Host-Client während dieser Bearbeitungssitzung nicht mehr gefragt. Administratoren können diesen Grenzwert mit dem  **AlertInterval** -Registrierungsschlüssel erhöhen, damit diese Warnmeldung nicht so oft angezeigt wird, wenn Benutzer prozessorlastige Add-Ins ausführen.
    
-  **Arbeitsspeicherauslastung** – Ein Grenzwert für die standardmäßige Arbeitsspeicherauslastung, der basierend auf dem verfügbaren physischen Arbeitsspeicher des Geräts dynamisch festgelegt wird.
    
    Wenn ein Host-Rich-Client feststellt, dass die Auslastung des physischen Arbeitsspeichers eines Geräts 80 % des verfügbaren Arbeitsspeichers übersteigt, beginnt der Client standardmäßig mit der Überwachung der Arbeitsspeicherauslastung durch das Add-In. Für Inhalts-Add-Ins und Add-Ins für den Aufgabenbereich erfolgt dies auf Dokumentenebene und für Outlook-Add-Ins auf Postfachebene. Der Client warnt den Benutzer standardmäßig in Zeitabständen von 5 Sekunden, wenn die Auslastung des physischen Arbeitsspeichers für eine Reihe von Apps auf Dokumenten- oder Postfachebene 50 % übersteigt. Dieser Grenzwert für die Arbeitsspeicherauslastung bezieht sich auf den physischen und nicht auf den virtuellen Arbeitsspeicher, um auf Geräten mit begrenztem Arbeitsspeicher wie etwa Tablets eine gute Leistung zu gewährleisten. Diese dynamische Einstellung kann vom Administrator durch einen festen Grenzwert ersetzt werden, indem der Windows-Registrierungsschlüssel  **MemoryAlertThreshold** als globale Einstellung in der Windows-Registrierung verwendet wird. Passen SIe dann das Warnintervall mithilfe des **AlertInterval** -Schlüssels als globale Einstellung an.
    
-  **Absturztoleranz** – Ein standardmäßiger Grenzwert von 4 Abstürzen für ein Add-In.
    
    Administratoren können den Grenzwert für Abstürze anpassen, indem sie die  **RestartManagerRetryLimit** -Registrierungsschlüssel verwenden.
    
-  **Nicht reagierende Anwendungen** – Grenzwert für den Zeitraum, in der ein Add-In nicht reagiert, von 5 Sekunden.
    
    Das hat Auswirkungen auf die Benutzerfreundlichkeit des Add-Ins und der Hostanwendung. Ist dies der Fall, startet die Hostanwendung automatisch alle aktiven Add-Ins für ein Dokument oder Postfach (falls zutreffend) neu und gibt unter Nennung des Add-Ins, das nicht mehr reagiert, eine Warnung an den Benutzer aus. Add-Ins können diesen Grenzwert erreichen, wenn sie bei der Durchführung von Aufgaben mit langer Bearbeitungszeit nicht regelmäßig Ausgaben erzeugen. Es gibt Verfahren, um sicherzustellen, dass die Reaktionsfähigkeit nicht beeinträchtigt wird. Dieser Grenzwert kann von Administratoren außer Kraft gesetzt werden.
    
     **Outlook-Add-Ins**
    
    Werden von einem Outlook-Add-In die o. g. Grenzwerte für die Auslastung der CPU-Kerne oder des Arbeitsspeichers bzw. die Toleranzgrenze für Abstürze überschritten, wird das Add-In von Outlook deaktiviert. Im Exchange Admin Center wird der Status der App als deaktiviert angezeigt.
    
     >**Hinweis**  Obwohl die Ressourcennutzung nur von Outlook-Rich-Clients und nicht von Outlook Web App oder OWA für mobile Geräte überwacht wird, wird ein Outlook-Add-In, die von einem Rich-Client deaktiviert wurde, auch für Outlook Web App und OWA für mobile Geräte deaktiviert.

    Zusätzlich zu den Bestimmungen bezüglich der CPU-Kerne, des Arbeitsspeichers und der Zuverlässigkeit sollten für Outlook-Add-Ins die folgenden Regeln für die Aktivierung eingehalten werden:
    
      -  **Antwortzeit für reguläre Ausdrücke** – Für Outlook gilt ein standardmäßiger Grenzwert von 1.000 Millisekunden für die Überprüfung aller regulären Ausdrücke im Manifest einem Outlook-Add-In. Bei einer Überschreitung des Grenzwerts versucht Outlook die Überprüfung zu einem späteren Zeitpunkt.
    
    Mithilfe einer Gruppenrichtlinie oder anwendungsspezifischen Einstellung in der Windows-Registrierung können Administratoren diesen standardmäßigen Grenzwert von 1.000 Millisekunden in der Einstellung  **OutlookActivationAlertThreshold** anpassen. Weitere Informationen finden Sie unter [Überschreiben von Ressourceneinsatzeinstellungen für die Leistung von Apps für Office](da14ec8c-5075-4035-a951-fc3c2b15c04b.md).
    
  -  **Erneute Überprüfung regulärer Ausdrücke** – Für Outlook gilt ein standardmäßiger Grenzwert von drei Versuchen für die erneute Überprüfung aller regulären Ausdrücke in einem Manifest. Schlägt die Überprüfung drei Mal fehl, weil der anwendbare Grenzwert (entweder die Standardeinstellung von 1.000 Millisekunden oder ein durch **OutlookActivationAlertThreshold** festgelegter Wert, wenn dieser in der Windows-Registrierung vorhanden ist) überschritten wurde, wird das Outlook-Add-In von Outlook deaktiviert. Im Exchange Admin Center wird der Status als deaktiviert angezeigt, und das Add-In wird für die Outlook-Rich-Clients, Outlook Web App und OWA für mobile Geräte deaktiviert.
    
    Mithilfe einer Gruppenrichtlinie oder anwendungsspezifischen Einstellung in der Windows-Registrierung können Administratoren die Anzahl der Versuche für die erneute Überprüfung in der Einstellung  **OutlookActivationManagerRetryLimit** anpassen. Weitere Informationen finden Sie unter [Überschreiben von Ressourceneinsatzeinstellungen für die Leistung von Apps für Office](da14ec8c-5075-4035-a951-fc3c2b15c04b.md).
    

     **Aufgabenbereichs- und Inhalts-Add-Ins**
    
    Werden von einem Inhalts-Add-In oder einem Add-In für den Aufgabenbereich die o. g. Grenzwerte für die Auslastung der CPU-Kerne oder des Arbeitsspeichers bzw. die Toleranzgrenze für Abstürze überschritten, wird dem Benutzer von der entsprechenden Hostanwendung eine Warnung angezeigt. Der Benutzer kann daraufhin:
    
      - Das Add-In neu starten.
    
  - Weitere Warnungen hinsichtlich der Überschreitung dieses Grenzwerts unterdrücken. Im Optimalfall löscht der Benutzer das Add-In aus dem Dokument, da eine weitere Verwendung des Add-Ins zusätzliche Leistungs- und Stabilitätsprobleme verursachen kann.
    

## Behandeln von Problemen mit der Ressourcennutzung im Telemetrieprotokoll


Office verfügt über ein Telemetrieprotokoll, in dem bestimmte Ereignisse von Office-Lösungen aufgezeichnet werden (Laden, Öffnen, Schließen und Fehler), die auf einem lokalen Computer ausgeführt werden, darunter Ressourcennutzungsprobleme in einer Office-Add-In. Wenn Sie das Telemetrieprotokoll eingerichtet haben, können Sie Excel zum Öffnen des Telemetrieprotokolls vom folgenden Standardspeicherort auf dem lokalen Laufwerk verwenden:

%Users%\ [Aktueller Benutzer]\AppData\Local\Microsoft\Office\15.0\Telemetry

Für jedes Ereignis, das vom Telemetrieprotokoll für ein Add-In nachverfolgt wird, wird Folgendes angegeben: Datum/Uhrzeit des Auftretens, Ereignis-ID, Schweregrad, kurzer aussagekräftiger Titel für das Ereignis, Anzeigename und eindeutige ID des Add-Ins sowie die Anwendung, von der das Ereignis protokolliert wurde. Sie können das Telemetrieprotokoll aktualisieren, um die derzeit nachverfolgten Ereignisse anzuzeigen. In der folgenden Tabelle sind einige Beispiele für Outlook-Add-Ins aufgeführt, die im Telemetrie-Protokoll nachverfolgt wurden. 



|**Datum/Uhrzeit**|**Ereignis-ID**|**Schweregrad**|**Titel**|**Datei**|**ID**|**Anwendung**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|08.10.12 17:57:10|7||Das Add-In-Manifest wurde erfolgreich heruntergeladen.|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|08.10.12 17:57:01|7||Das Add-In-Manifest wurde erfolgreich heruntergeladen.|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|
 In der folgenden Tabelle sind die Ereignisse aufgeführt, die im Telemetrieprotokoll für Office-Add-Ins normalerweise nachverfolgt werden.



|**Ereignis-ID**|**Titel**|**Schweregrad**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|7|Das Add-In-Manifest wurde erfolgreich heruntergeladen.||Das Manifest der Office-Add-In wurde von der Hostanwendung erfolgreich geladen und gelesen.|
|8|Das Add-In-Manifest konnte nicht heruntergeladen werden.|Kritisch|Die Hostanwendung konnte die Manifestdatei für die Office-Add-In nicht aus dem SharePoint-Katalog, dem Unternehmenskatalog oder aus dem Office Store laden.|
|9|Der Markup-Code des Add-Ins konnte nicht analysiert werden.|Kritisch|Die Hostanwendung hat das Office-Add-In-Manifest geladen, konnte jedoch das HTML-Markup der App nicht lesen.|
|10|Die CPU-Nutzung des Add-Ins ist zu hoch.|Kritisch|Die Office-Add-In hat über eine begrenzte Zeitspanne mehr als 90 % der CPU-Ressourcen beansprucht.|
|15|Das Add-In wurde aufgrund eines Timeouts bei der Zeichenfolgensuche deaktiviert.||Outlook-Add-Ins durchsuchen die Betreffzeile und die Nachricht einer E-Mail, um zu ermitteln, ob sie mithilfe eines regulären Ausdrucks angezeigt werden soll. Das Outlook-Add-In, die in der Spalte  **File** angegeben ist, wurde von Outlook deaktiviert, weil es bei dem Versuch, einen regulären Ausdruck zuzuordnen, mehrfach zu einer Zeitüberschreitung gekommen ist.|
|18|Das Add-In wurde erfolgreich geschlossen.||Die Hostanwendung konnte die Office-Add-In erfolgreich schließen.|
|19|Für das Add-In ist ein Laufzeitfehler aufgetreten, der lokal protokolliert wurde.|Kritisch|Bei der Office-Add-In ist ein Problem aufgetreten, das zu einem Fehler geführt hat. Weitere Einzelheiten können Sie dem Protokoll der  **Microsoft Office-Benachrichtigungen** entnehmen. Zeigen Sie das Protokoll mit der Windows-Ereignisanzeige auf dem Computer an, auf dem der Fehler aufgetreten ist.|
|20|Das Add-In konnte die Lizenzierung nicht überprüfen.|Kritisch|Die Lizenzinformationen für die Office-Add-In konnten nicht überprüft werden. Möglicherweise ist die Lizenz abgelaufen. Weitere Einzelheiten können Sie dem Protokoll der  **Microsoft Office-Benachrichtigungen** entnehmen. Zeigen Sie das Protokoll mit der Windows-Ereignisanzeige auf dem Computer an, auf dem der Fehler aufgetreten ist.|
Weitere Informationen finden Sie unter [Deploying Telemetry Dashboard](http://msdn.microsoft.com/de-de/library/f69cde72-689d-421f-99b8-c51676c77717%28Office.15%29.aspx) und [Problembehandlung für Office-Dateien und benutzerdefinierte Lösungen mit dem Telemetrieprotokoll](http://msdn.microsoft.com/library/ef88e30e-7537-488e-bc72-8da29810f7aa%28Office.15%29.aspx)


## Design- und Implementierungsverfahren


Die Grenzwerte für die Auslastung der CPU-Kerne und des Arbeitsspeichers sowie für die Absturztoleranz und die Reaktionsfähigkeit der Benutzeroberfläche gelten zwar nur für Office-Add-Ins, die auf den Rich-Clients ausgeführt werden. Trotzdem sollten Sie der optimalen Nutzung dieser Ressourcen und des Akkus höchste Priorität einräumen, wenn Ihr Add-In auf allen unterstützten Clients und Geräten gut laufen soll. Die Optimierung ist besonders dann wichtig, wenn Ihr Add-In Aufgaben mit langer Bearbeitungszeit durchführt oder große Datasets verarbeitet. In der folgenden Liste werden einige Verfahren vorgeschlagen, mit denen prozessorlastige oder datenintensive Vorgänge aufgeteilt werden können, damit Ihr Add-In nicht übermäßig Ressourcen in Anspruch nimmt und die Hostanwendung reaktionsfähig bleibt:


- In einem Szenario, in dem Ihr Add-In eine große Datenmenge aus einem unbegrenzten Dataset lesen muss, können Sie eine Auslagerung anwenden, wenn die Daten aus einer Tabelle gelesen werden, oder die Größe der Daten in jedem kürzeren Lesevorgang reduzieren, anstatt zu versuchen, den Lesevorgang in einem einzigen Prozess durchzuführen. 
    
    Unter [Wie kann ich bei einer intensiven JavaScript-Verarbeitung (kurzfristig) wieder die Steuerung über den Browser ermöglichen?](http://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript) finden Sie ein Codebeispiel für JavaScript und jQuery, in dem veranschaulicht wird, wie eine Reihe von möglicherweise langwierigen und prozessorlastigen Ein- und Ausgabevorgängen mit unbegrenzten Daten aufgeteilt wird. In diesem Beispiel wird die [setTimeout](http://msdn.microsoft.com/de-de/library/ie/ms536753%28v=vs.85%29.aspx)-Methode des globalen Objekts verwendet, um die Dauer der Ein- und Ausgabe zu beschränken. Des Weiteren werden die Daten in definierten Teilen verarbeitet, und nicht als zufällig unbegrenzte Daten.
    
- Wenn Ihr Add-In einen prozessorlastigen Algorithmus für die Verarbeitung eines großen Datenvolumens verwendet, können Sie Web-Worker verwenden, um die Aufgabe mit der langen Bearbeitungszeit im Hintergrund auszuführen, während ein separates Skript (z. B. zur Anzeige des Fortschritts auf der Benutzeroberfläche) im Vordergrund ausgeführt wird. Web-Worker blockieren keine Benutzerinteraktionen und sorgen dafür, dass die HTML-Seite reaktionsfähig bleibt. Ein Beispiel für Web-Worker finden Sie unter [Web Worker-Grundlagen](http://www.mdl5rocks.com/de/tutorials/workers/basics/). Weitere Informationen zu der Internet Explorer-Web-Worker-API finden Sie unter [Web-Worker](http://msdn.microsoft.com/de-de/library/IE/hh772807%28v=vs.85%29.aspx).
    
- Wenn Ihr Add-In einen prozessorlastigen Algorithmus verwendet, die Datenein- oder ausgabe jedoch in kleinere Teile aufgeteilt werden kann, können Sie einen Webdienst erstellen, die Daten an diesen übergeben, um den Prozessor zu entlasten, und auf einen asynchronen Rückruf warten.
    
- Testen Sie Ihre App immer bei dem höchsten zu erwartenden Datenaufkommen, und schränken Sie Ihr Add-In so ein, dass Verarbeitungen nur bis zu diesem Grenzwert möglich sind.
    

## Zusätzliche Ressourcen



- [Datenschutz und Sicherheit bei Office-Add-Ins](87c59a88-10e2-4c88-b6a8-736bd356e5f8.md)
    
- [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](e0c9e3d0-517e-4333-b8bd-e169c51a07f6.md)
    
