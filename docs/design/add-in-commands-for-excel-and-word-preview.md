
# Add-In-Befehle für Excel, Word und PowerPoint (Vorschau)
Verwenden Sie Add-In-Befehle in Ihrem Add-In, um die standardmäßige Office-Benutzeroberfläche in Excel, Word und PowerPoint zu erweitern.

 _**Gilt für:** apps for Office | Excel 2016 | Office Add-ins | PowerPoint | Word 2016_

Add-In-Befehle sind Benutzeroberflächenelemente, die die standardmäßige Office-Benutzeroberfläche erweitern und Aktionen in Ihrem Add-In starten. Sie können eine Schaltfläche auf dem Menüband oder ein Element zu einem Kontextmenü hinzufügen. Wenn Benutzer einen Add-In-Befehl auswählen, werden Aktionen, wie das Ausführen von JavaScript-Code oder das Anzeigen einer Add-In-Seite in einem Aufgabenbereich initiiert. Add-In-Befehle können Benutzer Ihr Add-In leichter finden und verwenden, wodurch die Akzeptanz und Wiederverwendung Ihres Add-Ins gesteigert und die Kundenbindung verbessert werden kann.

Add-In-Befehle können für Folgendes verwendet werden:


- Hinzufügen von Schaltflächen oder einer Dropdownliste von Schaltflächen zum Menüband.
    
- Hinzufügen einzelner Menüelemente mit optionalem Untermenü zum Kontextmenü.
    
- Anzeigen des Add-Ins, das eine Benutzeroberfläche für den Benutzer zur Interaktion anzeigt.
    
- Ausführen von JavaScript-Code, der normalerweise ohne Anzeigen einer Benutzeroberfläche angezeigt wird.
    
- Stellen Sie sicher, dass Ihre Add-Ins leichter zu finden sind, indem Sie mehrere Orte für Benutzer zum Öffnen des Add-Ins angeben. Sie können Ihr Add-In von Schaltflächen oder Kontextmenüs aus starten oder durch Verwendung von  **Einfügen** > **Meine Add-Ins**.
    
- Öffnen unterschiedlicher Seiten in Ihrem Add-In. Wenn bisher ein Add-In geöffnet wurde, wurde es immer auf der gleichen Seite gestartet. Um zu anderen Seiten in Ihrem Add-In überzugehen, musste der Code die Benutzeroberflächenübergange verwalten. Mit Add-In-Befehlen, können Sie eine beliebige Seite in Ihrem Add-In öffnen.
    
- Anzeigen mehrerer Aufgabenbereich-Add-Ins gleichzeitig in Excel 2016 oder Word 2016.
    
Eine Übersicht über Add-In-Befehle in Excel und Word finden Sie unter [Add-In-Befehle im Office-Menüband (Öffentliche Vorschau)](https://channel9.msdn.com/Events/Visual-Studio/Connect-event-2015/316).


## Erste Schritte beim Erstellen von Add-In-Befehlen

 **Einstieg in die Verwendung von Add-In-Befehlen**


1. Laden Sie das [Beispiel für Office-Add-In-Befehle](https://github.com/OfficeDev/Office-Add-in-Command-Sample) herunter.
    
2. Erstellen Sie Ihren Add-In-Befehl, indem Sie Ihr Manifest, wie in [Erstellen von Add-In-Befehlen in Ihrem Manifest für Excel, Word und PowerPoint (Vorschau)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md) beschrieben, aktualisieren. Weitere Informationen finden Sie unter [Bewährte Methoden für die Entwicklung von Office-Add-Ins](../../docs/design/add-in-development-best-practices.md).
    

## Installieren und Testen Ihres Add-Ins

Nachdem Sie Ihre Manifestdatei aktualisiert haben, installieren Sie Ihr Add-In, und testen Sie Ihre Add-In-Befehle. Sie können Ihre Add-In-Befehle testen auf:


- Excel Online
    
- Office auf einem Windows Desktop
    

### Installieren Ihres Add-Ins auf Excel Online



1. Öffnen Sie [Microsoft Office Online](https://office.live.com/).
    
2. Wählen Sie unter  **Jetzt mit den Online-Apps beginnen** die Option **Excel** > **Neue leere Arbeitsmappe** > **Einfügen** aus.
    
3. Wählen Sie unter  **Add-Ins** die Option **Office Add-Ins**.
    
4. Wählen Sie unter  **Office Add-Ins** die Option **Mein Add-In hochladen** > **Mehr** ( **...**) aus.
    
5. Wählen Sie die Manifestdatei auf Ihrem Computer aus, und wählen Sie dann  **Add-In hochladen** aus.
    
6. Überprüfen Sie, ob Ihre Add-In-Befehle auf dem Menüband oder im Kontextmenü angezeigt werden.
    

### Bekannte Probleme bei Add-In-Befehlen in Excel Online



- Wenn Sie Excel Online schließen, werden auch die Befehle Add-In und Add-In geschlossen. Wenn Sie ein Dokument in Excel Online das nächste Mal öffnen, müssen Sie die in   [Installieren Ihres Add-Ins auf Excel Online](#installieren-ihres-add-ins-auf-excel-online) beschriebenen Schritte erneut ausführen.
    
- Wenn Sie Ihre Add-In-Befehle in Ihrem Add-In verwenden möchten, laden Sie das Add-In über  **Einfügen** > **Office Add-Ins** > **Mein Add-In hochladen** hoch. Add-Ins, die an den Office Store oder an einen Add-In-Katalog einer internen Organisation in SharePoint übermittelt wurden, zeigen keine Add-In-Befehle an.
    
- Nur Excel Online unterstützt Add-In-Befehle. Unterstützung für Word Online und PowerPoint Online wird demnächst hinzugefügt.
    
- Wenn Sie  **Einfügen** > **Office Add-Ins** > **MEINE ORGANISATION** auswählen und zum Office Store umgeleitet werden, versuchen Sie, sich bei Excel Online anzumelden. Wählen Sie dann erneut **MEINE ORGANISATION** aus.
    
- Wenn der Aufgabenbereich nicht aktualisiert wird, nachdem Sie unterschiedliche Add-In-Befehle ausgewählt haben, aktualisieren Sie die Seite in Ihrem Browser.
    

### Installieren Ihres Add-Ins in Excel 2016 oder Word 2016 auf einem Windows Desktop


Bevor Sie das Add-In installieren, sollten Sie überprüfen, ob Sie mindestens Excel 2016 oder Word 2016, Version 16.0.6326, ausführen.


1. Öffnen Sie  **Excel 2016** > **Leere Arbeitsmappe** > **Datei** > **Konto**.
    
2. Stellen Sie unter  **Office Updates** sicher, dass die Versionsnummer mindestens **16.0.6326** ist.
    
Um die neueste Version von Office zu erhalten, die Entwicklerfeatures in der Vorschau enthält, führen Sie in Abhängigkeit von Ihrem Abonnement einen der folgenden Schritte aus:


- Wenn Sie ein Abonnent von Office 365-Home, Office 365 Personal oder Office 365-University sind, installieren Sie den [Office-Insider-Build](https://products.office.com/de-de/office-insider).
    
    oder
    
- Wenn Sie ein kommerzieller Office 365-Abonnent sind, wählen Sie [First Release](https://support.office.com/de-de/article/Office-365-release-options-3B3ADFA4-1777-4FF0-B606-FB8732101F47?ui=en-US&amp;rs=en-001&amp;ad=DE) aus, und führen Sie dann die folgenden Schritte aus.
    

1. Laden Sie das [Office 2016-Bereitstellungstools](http://www.microsoft.com/en-us/download/details.aspx?id=49117) herunter, und führen Sie es aus.
    
2. Ersetzen Sie  **configuration.xml** durch die [First Release-Konfigurationsdatei](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
    
3. Öffnen Sie ein Eingabeaufforderungsfenster mit erhöhten Rechten (Als Administrator ausführen), und führen Sie den folgenden Befehl aus.
    
  ```
  setup.exe /configuration configuration.xml
  ```

4. Überprüfen Sie, ob die Versionsnummer von Office gleich  **16.0.6326.0000** oder höher für Excel oder Word und **16.0.6568.2025** oder höher für PowerPoint ist.
    
Bevor Sie die folgenden Schritte ausführen, stellen Sie sicher, dass Ihr Add-In ohne Fehler ausgeführt wird. Wenn Sie zum Beispiel Visual Studio verwenden, drücken Sie F5, um zu überprüfen, dass das Add-In ohne Fehler ausgeführt wird. Wenn das Add-In nach dem Durchführen der folgenden Schritte Fehler generiert, können Sie die Fehlerquelle leichter identifizieren.

 **So installieren Sie Ihr Add-In auf Office auf einem Windows Desktop**


1. Laden Sie den Registrierungsschlüssel aus den [Beispielen für Office Add-In-Befehle](https://github.com/OfficeDev/Office-Add-in-Command-Sample) herunter, und führen Sie ihn aus, um Add-In-Befehle in Office zu aktivieren.
    
2. Veröffentlichen Sie Ihr Add-In-Manifestdatei in einer Netzwerkdateifreigabe. Weitere Informationen finden Sie unter [Erstellen eines Katalogs mit freigegebenen Ordnern für Aufgabenbereich- und Inhalts-Add-ins](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    
3. Schließen Sie die Office-Hostanwendung (Excel 2016 oder Word 2016), und öffnen Sie sie erneut.
    
4. Wählen Sie  **Einfügen** > **Meine Add-Ins** > **FREIGEGEBENER ORDNER** aus.
    
5. Wählen Sie Ihr Add-In aus, und wählen Sie dann  **OK** aus.
    
6. Überprüfen Sie, ob Ihre Add-In-Befehle in Office angezeigt werden. Wenn Ihre Add-In-Befehle nicht in der Office-Benutzeroberfläche angezeigt werden, wählen Sie  **Einfügen** > **Meine Add-Ins** > **Aktualisieren** aus.
    

### Bekannte Probleme bei Add-In-Befehlen in Office auf Windows Desktop



- Wenn Sie eine Klick-und-Los-Version von Office, z. B. Office 365 ProPlus, verwenden, stellen Sie sicher, dass die Installation der Updates abgeschlossen ist, bevor Sie Ihr Add-In installieren.
    
- Add-In-Befehle werden nur in Excel 2016, Word 2016 und PowerPoint 2016 unterstützt.
    
- Add-Ins, die an den Office Store oder an einen Add-In-Katalog einer internen Organisation in SharePoint übermittelt wurden, zeigen keine Add-In-Befehle an. Sie können Ihre Manifestdatei nur von der Dateifreigabe laden.
    
- Aktualisieren an Add-In-Befehlen werden in Office nur gerendert, wenn Sie eine Aktualisierung ausführen. Wenn eine der folgenden Situationen eintritt, wählen Sie  **Einfügen** > **Meine Add-Ins** > **Aktualisieren** aus:
    
      - Sie möchten nicht funktionierende oder doppelte Schaltflächen aus dem Menüband entfernen. Vielleicht versuchen Sie auch, den Cache zu löschen, indem Sie alle Dateien unter  **%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\AppCommands** löschen und dann **Einfügen** > **Meine Add-Ins** > **Aktualisieren** auswählen.
    
  - Das Aufgabenbereich-Add-In wird nicht geladen.
    
  - Sie müssen Ihre Add-In-Befehle aktualisieren. Damit Ihre Add-In-Befehle auf dem Menüband angezeigt werden, müssen Sie möglicherweise zwei Mal aktualisieren.
    
- Wenn Sie Ihre Symboldateien nach der Installation des Add-Ins ändern möchten, führen Sie einen der folgenden Schritte aus:
    
      - Gehen Sie zu  **%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\icons**, und löschen Sie alle Dateien in diesem Ordner.
    
    oder
    
  - Ändern Sie den Dateinamen des Symbols im **&lt;Resources&gt;**-Element im Manifest, und veröffentlichen Sie Ihr Add-In erneut.
    
- Add-In-Befehle stehen immer in der Office-Benutzeroberfläche zur Verfügung, auch wenn kein Dokument angezeigt wird.
    

## Zusätzliche Ressourcen



- [Erstellen von Add-In-Befehlen in Ihrem Manifest für Excel, Word und PowerPoint (Vorschau)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md)
    
- [Bewährte Methoden für die Entwicklung von Office-Add-Ins](../../docs/design/add-in-development-best-practices.md)
    
