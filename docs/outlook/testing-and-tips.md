
# Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken
In diesem Thema erfahren Sie, wie Sie ein Outlook-Add-In zu Testzwecken im Rahmen des Outlook-Add-In-Entwicklungszyklus bereitstellen und installieren. 

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Bereitstellen eines Outlook-Add-Ins zu Testzwecken


Im Rahmen des Prozesses zum Entwickeln eines Outlook-Add-Ins werden Sie wahrscheinlich das Add-In wiederholt zu Testzwecken bereitstellen und installieren, was die folgenden Schritte beinhaltet:


1. Erstellen einer Manifestdatei zur Beschreibung des Add-Ins
    
2. Bereitstellen der Add-In-Benutzeroberflächendatei(en) auf einem Webserver
    
3. Installieren des Add-Ins in Ihrem Postfach
    
4. Testen des Add-Ins, Vornehmen entsprechender Änderungen an den Benutzeroberflächen- oder Manifestdateien und Wiederholen der Schritte 2 und 3 zum Testen der Änderungen
    

### Erstellen einer Manifestdatei für das Add-In

Jedes Add-In wird durch ein XML-Manifest beschrieben. Hierbei handelt es sich um ein Dokument, das die Server-Informationen zum Add-In liefert, beschreibende Informationen zum Add-In für den Benutzer enthält sowie den Speicherort der HTML-Datei für die Add-In-Benutzeroberfläche identifiziert. Sie können das Manifest in einem lokalen Ordner oder auf einem lokalen Server speichern, vorausgesetzt, dass auf den Exchange-Server des Postfachs, das Sie zum Testen verwenden, zugegriffen werden kann. Wir gehen davon aus, dass Sie Ihr Manifest in einem lokalen Ordner speichern. Informationen zum Erstellen einer Manifestdatei finden Sie unter [Outlook-Add-In-Manifeste](98ee2cca-f866-4377-bd1f-6a5338ebd7b1.md). 


### Bereitstellen eines Add-Ins auf einem Webserver

Die Add-In-Benutzeroberfläche können Sie mithilfe von HTML und JavaScript erstellen. Die resultierende Quelldatei wird auf einem Webserver gespeichert, auf den der Exchange-Server Zugriff hat, von dem das Add-In gehostet wird. Die Quelldatei wird durch das untergeordnete Element  **SourceLocation** im Element [DesktopSettings](http://msdn.microsoft.com/de-de/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx), [TabletSettings](http://msdn.microsoft.com/de-de/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) und/oder [PhoneSettings](http://msdn.microsoft.com/de-de/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) identifiziert, das in der Add-In-Manifestdatei angegeben ist..

Nach der Bereitstellung der Benutzeroberflächendateien für das Add-In können Sie die Add-In-Benutzeroberfläche und das Verhalten aktualisieren, indem Sie die auf dem Webserver gespeicherte HTML-Datei durch eine neue Version der HTML-Datei ersetzen.


## Installieren des Add-Ins


Nachdem Sie die Add-In-Manifestdatei vorbereitet und die Add-in-Benutzeroberfläche auf einem Webserver bereitgestellt haben, auf den der Zugriff möglich ist, können Sie das Add-In für ein Postfach auf einem Exchange-Server installieren. Hierzu können Sie einen Outlook-Rich-Client, Outlook Web App oder OWA für mobile Geräte verwenden oder aber Windows PowerShell-Remote-Cmdlets ausführen.


### Installieren eines Add-Ins in einem Outlook-Rich-Client

Sie können ein Add-In installieren, wenn sich Ihr Postfach auf Exchange Online, Exchange 2013 oder einer späteren Version befindet. In Outlook für Windows können Sie Add-Ins über die Office Fluent Backstage-Ansicht installieren. Wählen Sie  **Datei** und **Add-Ins verwalten** aus. Damit können Sie sich beim Exchange Admin Center anmelden. Nach der Anmeldung setzen Sie den Installationsvorgang mit Schritt 4 im nächsten Abschnitt fort.

In Outlook für Mac wählen Sie  **Add-ins verwalten** am rechten Ende der Add-in-Leiste aus und melden sich dann beim Exchange Admin Center an. Fahren Sie mit Schritt 4 im nächsten Abschnitt fort.


### Installieren eines Add-Ins mithilfe der Outlook Web App oder mithilfe von Outlook.com

Führen Sie die folgenden Schritte aus, um ein Outlook-Add-In mit der Outlook Web App (OWA) zu installieren:


1. Navigieren Sie zu der OWA-URL für Ihre Organisation oder zu Outlook.com, und melden Sie sich an.
    
2. Klicken Sie auf das Zahnradsymbol in der oberen rechten Ecke, und wählen Sie  **Add-Ins verwalten** aus.
    
3. 4. Wählen Sie das Pluszeichen ( **+**) aus, um ein neues Add-in hinzuzufügen.
    
5. Wählen Sie in der Dropdownliste die Option  **Aus Datei hinzufügen**, falls Sie das Manifest in einem lokalen Ordner gespeichert haben.
    
6. Wechseln Sie zum Dateipfad des Manifests, und wählen Sie  **Installieren** aus.
    
7. Wählen Sie den Benutzernamen in der rechten oberen Ecke des Fensters aus, und wählen Sie dann  **Meine E-Mail** aus, um zu Ihrem E-Mail-Programm zu wechseln und das Add-In zu testen.
    

(../../images  Wenn Sie nicht eine der folgenden Anwendungen zum Entwickeln Ihres Add-Ins verwenden: Wenn Sie auch nicht mindestens über die Rolle „Eigene benutzerdefinierte Add-Ins" für Exchange Server verfügen, können Sie Add-Ins nur über den Office Store installieren. Um das Add-In zu testen oder Add-Ins generell durch Angeben einer URL oder eines Dateinamens für das Add-In-Manifest zu installieren, sollten Sie Ihren Exchange-Administrator bitten, die erforderlichen Berechtigungen bereitzustellen.Der Exchange-Administrator kann das folgende PowerShell-Cmdlet ausführen, um einem einzelnen Benutzer die erforderlichen Berechtigungen zuzuweisen. In diesem Beispiel ist wendyri der E-Mail-Alias des Benutzers.New-ManagementRoleAssignment –Role "My Custom add-ins" –User "wendyri"Bei Bedarf kann der Administrator das folgende Cmdlet ausführen, um mehreren Benutzern ähnliche erforderliche Berechtigungen zuzuweisen:$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom add-ins" -User $_.Alias}Weitere Informationen zur Rolle „Eigene benutzerdefinierte Add-ins" finden Sie unter [Rolle „Eigene benutzerdefinierte Add-ins"](http://technet.microsoft.com/de-de/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28exchg.150%29.aspx).Wenn Sie Office 365, Napa oder Visual Studio zum Entwickeln von Add-Ins verwenden, wird Ihnen die Organisationsadministrator-Rolle zugewiesen. Damit können Sie Add-Ins mithilfe einer Datei oder URL im EAC installieren, oder Sie können Powershell-Cmdlets verwenden.


### Installieren eines Add-Ins mit der Exchange-Verwaltungskonsole

Nachdem Sie auf Ihrem Exchange-Server eine Windows PowerShell-Remotesitzung eingerichtet haben, können Sie mit dem Cmdlet  **New-App** ein Outlook-Add-In mithilfe des folgenden PowerShell-Befehls installieren.


```
New-App –URL:"http://<fully-qualified URL">
```

Die vollqualifizierte URL ist der Speicherort der Add-in-Manifestdatei, die Sie für Ihr Add-in vorbereitet haben.

Sie können die folgenden zusätzlichen PowerShell-Cmdlets zum Verwalten der Add-ins für ein Postfach verwenden:


-  **Get-App** – Listet die für ein Postfach aktivierten Add-Ins auf.
    
-  **Set-App** – Aktiviert oder deaktiviert ein Add-In in einem Postfach.
    
-  **Remove-App** – Entfernt ein zuvor installiertes Add-In von einem Exchange-Server.
    

## Inhalt dieses Abschnitts



- [Tipps zum Verwenden von Datumswerten in Outlook-Add-Ins](03366046-778c-4a34-8f85-b1e538d59c89.md)
    
- [Grenzwerte für Aktivierung und JavaScript-API für Outlook-Add-Ins](e0c9e3d0-517e-4333-b8bd-e169c51a07f6.md)
    
- [Problembehandlung für die Aktivierung von Outlook-Add-Ins](da5b56c9-7fd1-4556-8c0e-f489c4c9e9b6.md)
    

## Weitere Ressourcen



- [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md)
    
- [Problembehandlung von Benutzerfehlern mit Office Add-Ins](4e3a5129-5bb8-4aed-b122-200c6410d82c.md)
    
