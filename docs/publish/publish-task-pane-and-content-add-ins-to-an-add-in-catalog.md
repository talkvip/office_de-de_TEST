
# Veröffentlichen von Aufgabenbereich- und Inhalts-Add-Ins in einem Add-In-Katalog in SharePoint
Nachdem Sie Ihr Add-In-Manifest in einen Katalog hochgeladen haben, können Benutzer ihre Office-Clients so konfigurieren, dass Ihr Add-In über das Dialogfeld Office-Add-Ins verfügbar ist.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | PowerPoint | Project | Word_

Führen Sie die folgenden Schritte aus, um das Manifest für Ihr Aufgabenbereich- oder Inhalts-Add-in in einen Office-Add-Ins-Katalog in SharePoint hochzuladen. Informationen zum Einrichten eines Add-in-Katalogs finden Sie unter [Einrichten eines Add-In-Katalogs unter SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md).

 >**Hinweis**  Der Name "Apps für Office" wird in "Office-Add-Ins" geändert. Während dieses Übergangs kann in der Dokumentation und in der Benutzeroberfläche einiger Office-Anwendungen und Visual Studio-Tools möglicherweise noch der Begriff "App/Apps" verwendet werden. Weitere Informationen finden Sie unter [Neuer Name für Apps für Office und SharePoint: Office- und SharePoint-Add-Ins](https://msdn.microsoft.com/de-de/library/fp161507.aspx#Anchor_2).


## Veröffentlichen in einem Add-In-Katalog


1. Navigieren Sie zum Add-In-Katalog.
    
2. Wählen Sie den Link  **Zum Hinzufügen eines neuen Elements klicken** aus.
    
3. Wählen Sie  **Durchsuchen** aus, und geben Sie das hochzuladende Manifest an.
    
Nachdem Sie Add-In-Manifeste in den Office-Add-Ins-Katalog hochgeladen haben, können Benutzer ihre Installationen von Word, Excel, PowerPoint oder Project manuell konfigurieren, indem sie über  **Datei** > **Optionen** > **Trust Center** > **Einstellungen für das Trust Center** > **Vertrauenswürdige Add-In-Kataloge** die URL der _übergeordneten SharePoint-Websitesammlung_ dieses Add-In-Katalogs angeben. Hat die URL des Office-Add-Ins-Katalogs z. B. das Format:

 **https://**_Domäne_**/sites/**_AddInKatalogWebsitesammlung_**/AgaveCatalog**

…sollten Benutzer nur die URL der übergeordneten Websitesammlung angeben:

 **https://**_Domäne_**/sites/**_AddInKatalogWebsitesammlung_

Nachdem Sie die URL des Add-In-Katalogs angegeben haben, müssen Sie die Office-Anwendung schließen und erneut öffnen, damit der Add-In-Katalog im Dialogfeld  **OfficeAdd-Ins** verfügbar wird.

Alternativ dazu können Administratoren einen Office-Add-In-Katalog unter SharePoint mithilfe der Gruppenrichtlinie angeben. Dies ist im Abschnitt [Verwenden der Gruppenrichtlinie zum Verwalten der Installation und Nutzung von Office-Add-Ins durch Benutzer](https://technet.microsoft.com/de-de/library/jj219429.aspx) in [Übersicht über Office-Add-Ins 2013](https://technet.microsoft.com/de-de/library/jj219429.aspx) auf TechNet beschrieben.

Die Inhalts- und Aufgabenbereich-Add-Ins in diesem Katalog sind nun über das Dialogfeld  **Office-Add-Ins** verfügbar. Wählen Sie auf der Registerkarte **Einfügen** die Option **Meine Add-Ins** und dann die Option **MEINE ORGANISATION** aus.


## Suchen nach dem Add-In-Katalog

Falls bereits ein Add-In-Katalog für eine SharePoint-Webanwendung eingerichtet ist, können Sie mit den folgenden Schritten danach suchen:


1. Öffnen Sie die Hauptseite der SharePoint-Zentraladministration.
    
2. Wählen Sie  **Add-ins** aus.
    
3. Wählen Sie  **Add-In-Katalog verwalten** aus.
    
4. Wählen Sie den angezeigten Link und dann in der linken Navigationsleiste  **Office-Add-Ins** aus.
    

## Zusätzliche Ressourcen


- [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)
    
- [Einrichten eines Add-In-Katalogs unter SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md)
    
- [Einrichten eines Add-In-Katalogs in Office 365](../../publish/set-up-an-add-in-catalog-on-office-365.md)
    
