
# Problembehandlung von Benutzerfehlern mit Office Add-Ins
Behandeln von Problemen, die Ihre Benutzer mit Office-Add-ins haben.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Gelegentlich treten bei Benutzern Probleme mit Office-Add-Ins auf, die Sie entwickeln. Beispielsweise tritt bei einem Add-In ein Fehler auf oder es kann nicht darauf zugegriffen werden. Verwenden SIe die Informationen in diesem Artikel, um häufige Probleme zu beheben, die Benutzer mit Ihrem Office Add-In haben.

Sie können auch [Fiddler](http://www.telerik.com/fiddler) verwenden, um die Probleme mit Ihren Add-Ins zu ermitteln und zu debuggen.

Nachdem Sie das Problem des Benutzers gelöst haben, können Sie [direkt auf Kundenbewertungen im Office Store reagieren](https://msdn.microsoft.com/library/jj635874.aspx).

## Häufige Fehler und Schritte zur Problembehandlung

In der folgenden Tabelle sind die häufigsten Fehlermeldungen aufgeführt, die bei Benutzern auftreten können und die Schritte, mit denen die Benutzer die Fehler beheben können.



|**Fehlermeldung**|**Lösung**|
|:-----|:-----|
|App-Fehler: Katalog konnte nicht erreicht werden|Überprüfen Sie die Firewall-Einstellungen."Katalog" bezieht sich auf den Office Store. Diese Meldung deutet darauf hin, dass der Benutzer nicht auf den Office Store zugreifen kann.|
|APP-FEHLER: Diese App konnte nicht gestartet werden. Schließen Sie dieses Dialogfeld, um das Problem zu ignorieren oder klicken SIe auf "Neu starten", um es erneut zu versuchen.|Stellen Sie sicher, dass Sie die neuen Office-Updates installiert haben oder laden Sie das [Update für Office 2013](https://support.microsoft.com/de-de/kb/2986156/) herunter.|
|Fehler: Objekt unterstützt die Eigenschaft oder Methode 'defineProperty' nicht|Stellen Sie sicher, dass Internet Explorer nicht im Kompatibilitätsmodus ausgeführt wird. Wechseln Sie zu Extras >  **Compatibilitätsansichtseinstellungen**.|
|Es tut uns leid, die App konnte nicht geladen werden, da Ihre Browserversion nicht unterstützt wird. Klicken Sie hier, um eine Liste unterstützter Browserversionen anzuzeigen.|Stellen Sie sicher, dass der Browser den lokalen HTML5-Speicher unterstützt oder setzen Sie Ihre Internet Explorer-Einstellungen zurück.Weitere Informationen zu unterstützten Browsern finden Sie unter [Voraussetzungen zum Ausführen von Office-Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).|

## Outlook-Add-In funktioniert nicht korrekt

Wenn ein Outlook-Add-In unter Windows nicht korrekt funktioniert, aktivieren Sie das Skriptdebugging in Internet Explorer.


- Wechseln Sie zu Extras >  **Internetoptionen** > **Erweitert**.
    
- Deaktivieren Sie unter  **Browsen** das Kontrollkästchen **Skriptdebugging deaktivieren (Internet Explorer)** und **Skriptdebugging deaktivieren (Sonstige)**.
    
Wir empfehlen, diese Einstellungen nur für die Fehlerbehandlung zu deaktivieren. Wenn Sie sie deaktiviert lassen, erhalten Sie beim Browsen Eingabeaufforderungen. Aktivieren Sie nach der Behebung des Fehlers die Kontrollkästchen  **Skriptdebugging deaktivieren (Internet Explorer)** und **Skriptdebugging deaktivieren (Sonstige)** erneut.


## Add-In lässt sich in Office 2013 nicht aktivieren

Wenn sich das Add-In nicht aktivieren lässt, wenn der Benutzer folgende Schritte durchführt:


1. Er sich mit seinem Microsoft-Konto in Office 2013 anmeldet
    
2. die Überprüfung in zwei Schritten für sein Microsoft-Konto aktiviert
    
3. Er die Identität überprüft, wenn er dazu aufgefordert wird, bei dem Versuch ein Add-In einzufügen
    
Stellen Sie sicher, dass die aktuellsten Office-Updates installiert sind oder laden Sie das [Update für Office 2013](https://support.microsoft.com/de-de/kb/2986156/) herunter.


## Zusätzliche Ressourcen



- [Debuggen von Add-Ins in Office Online](../testing/debug-add-ins-in-office-online.md)
    
- [Querladen eines Office-Add-Ins auf einem iPad und einem Mac-Computer](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Debuggen von Office-Add-Ins auf dem iPad und einem Mac-Computer](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [Erstellen und Debuggen von Office-Add-Ins in Visual Studio](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md)
    
