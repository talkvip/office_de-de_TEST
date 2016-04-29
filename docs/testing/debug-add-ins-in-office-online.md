
# Debuggen von Add-Ins in Office Online
Verwenden Sie Office Online zum Debuggen von Add-Ins, ohne Office auf der Windows-Plattform auszuführen. z. B. bei Entwicklungen auf einem Mac. Sie.

 _**Gilt für:** apps for Office | Excel | Office Add-ins | Word_

Sie können Add-Ins auf einem Computer erstellen und debuggen, auf dem Windows oder der Office 2013- oder Office 2016-Desktopclient nicht ausführt wird. In diesem Artikel wird beschrieben, wie Sie Add-Ins mithilfe von Office Online testen und debuggen können. 

Erste Schritte:


- Besorgen Sie sich ein Office 365-Entwicklerkonto, sofern Sie noch keins besitzen oder Zugriff auf eine SharePoint-Website haben.
    
     >**Hinweis**  Um sich für ein kostenloses Office 365-Entwicklerkonto zu registrieren, treten Sie unserem [Office 365-Entwicklerprogramm](https://dev.office.com/devprogram) bei.
- Richten Sie einen Add-In-Katalog in Office 365 (SharePoint Online) ein. Ein Add-In-Katalog ist eine dedizierte Websitesammlung in SharePoint Online, in der Dokumentbibliotheken für Office-Add-Ins gehostet werden. Wenn Sie über eine eigene SharePoint-Website verfügen, können Sie eine Add-In-Katalogdokumentbibliothek einrichten. Diesbezügliche Informationen finden Sie unter:
    
      - [Einrichten eines Add-In-Katalogs in SharePoint Online](https://msdn.microsoft.com/DE-DE/library/office/dn574752.aspx)
    
  - [Einrichten eines Add-In-Katalogs unter SharePoint](https://msdn.microsoft.com/DE-DE/library/office/fp123530.aspx)
    

## Debuggen des Add-Ins aus Excel Online oder Word Online

So debuggen Sie das Add-In mithilfe vonOffice Online:


1. Stellen Sie das Add-In auf einem Server bereit, der SSL unterstützt.
    
     >**Hinweis**  Wir empfehlen die Verwendung des [Yeoman-Generators ](https://github.com/OfficeDev/generator-office) zum Erstellen und Hosten des Add-Ins.
2. Aktualisieren Sie in der [Add-In-Manifestdatei](../../docs/overview/add-in-manifests.md) den Wert des **SourceLocation**-Elements so, dass er einen absoluten anstelle eines relativen URI enthält. Beispiel:
    
     ` <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />`
    
3. Laden Sie das Manifest in die Office-Add-Ins-Bibliothek im Add-In-Katalog auf SharePoint hoch.
    
4. Starten Sie Excel Online oder Word Online über das App-Startfeld in Office 365, und öffnen Sie ein neues Dokument.
    
5. Wählen Sie auf der Registerkarte "Einfügen"  **Meine Add-Ins** oder **Office-Add-Ins** aus, um Ihr Add-In einzufügen und in der App zu testen.
    
6. Debuggen Sie das Add-In mit dem Tool-Debugger Ihres bevorzugten Browsers.
    
    Nachfolgend sind einige Probleme aufgefüht, die beim Debuggen auftreten können:
    
      - Einige JavaScript-Fehler, die angezeigt werden, stammen möglicherweise aus Office Online.
    
  - Der Browser zeigt möglicherweise einen Fehler zu einem ungültigen Zertifikat an, den Sie umgehen müssen.
    
  - Wenn Sie im Code Haltepunkte setzen, löst Office Online möglicherweise einen Fehler aus, der angibt, dass nicht gespeichert werden kann.
    

## Zusätzliche Ressourcen



- [Bewährte Designmethoden für Office Add-Ins](d455b76b-4d76-493d-a681-6b02ba1f38a8.md)
    
- [Validierungsrichtlinien für an den Office Store übermittelte Apps und Add-Ins (Version 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [Erstellen effektiver Office Store-Apps und -Add-Ins](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [Problembehandlung von Benutzerfehlern mit Office Add-Ins](../../docs/testing/testing-and-troubleshooting.md)
    
