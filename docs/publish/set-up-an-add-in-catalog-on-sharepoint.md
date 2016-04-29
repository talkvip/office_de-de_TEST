
# Einrichten eines Add-In-Katalogs unter SharePoint
Einrichten eines Add-In-Katalogs unter SharePoint zum Veröffentlichen von Aufgabenbereich- und Inhalts-Add-Ins

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | PowerPoint | Project | Word_

Ein Add-In-Katalog ist eine Dokumentbibliothek unter SharePoint, über die Manifestdateien für Aufgabenbereich- und Inhalts-Office-Add-Ins sowie SharePoint-Add-Ins veröffentlicht werden können. Bei Office-Add-Ins lädt ein Administrator eine [Manifestdatei](http://msdn.microsoft.com/de-de/library/4139ff24-afac-472a-af7d-9d069587ac9b%28Office.15%29.aspx) in den Add-In-Katalog hoch. Wenn ein Administrator einen Add-In-Katalog als vertrauenswürdig registriert (durch Festlegen einer Gruppenrichtlinie oder über die Seite **Vertrauenswürdiger Add-In-Katalog** des Dialogfelds **Optionen**, und zwar über  **Datei** > **Optionen** > **Sicherheitscenter** > **Einstellungen für das Sicherheitscenter** > **Vertrauenswürdige Add-In-Kataloge**), können Benutzer das Add-In über den Einfügebereich einer Office-Clientanwendung einfügen.

 >**Hinweis**  Der Name "Apps für Office" wird in "Office-Add-Ins" geändert. Während dieses Übergangs kann in der Dokumentation und in der Benutzeroberfläche einiger Office-Anwendungen und Visual Studio-Tools möglicherweise noch der Begriff "App/Apps" verwendet werden. Weitere Informationen finden Sie unter [Neuer Name für Apps für Office und SharePoint: Office- und SharePoint-Add-Ins](https://msdn.microsoft.com/de-de/library/fp161507.aspx#Anchor_2).

Pro SharePoint-Webanwendung kann nur ein Add-In-Katalog für Office-Add-Ins verwendet werden. So richten Sie den Add-In-Katalog für eine Webanwendung ein:

1. Wechseln Sie auf die  **Zentraladministrationswebsite** ( **Start** > **Alle Programme** > **Microsoft SharePoint 2013-Produkte** > **SharePoint 2013-Zentraladministration**).
    
2. Wählen Sie im linken Aufgabenbereich den Link  **Add-Ins** aus.
    
3. Wählen Sie auf der Seite  **Add-Ins** unter **Add-In-Verwaltung** die Option **Add-In-Katalog verwalten** aus.
    
4. Stellen Sie auf der Seite  **Add-In-Katalog verwalten** sicher, dass in der **Webanwendungsauswahl** die richtige Webanwendung ausgewählt ist.
    
5. Wählen Sie  **Websiteinstellungen anzeigen** aus.
    
6. Wählen Sie auf der Seite  **Websiteeinstellungen** die Option **Websitesammlungsadministratoren** aus, um die Websitesammlungsadministratoren anzugeben, und wählen Sie dann **OK** aus.
    
7. Um Benutzern Websiteberechtigungen zu gewähren, wählen Sie  **Websiteberechtigungen** und dann **Berechtigungen erteilen** aus.
    
8. Geben Sie im Dialogfeld  **Share ‘App Catalog Site’** einen oder mehrere Websitebenutzer an, legen Sie für diese die entsprechenden Berechtigungen fest, legen Sie optional andere Optionen fest, und wählen Sie dann **Freigeben** aus.
    
9. Wählen Sie zum Hinzufügen von Add-Ins zum Office-Add-Ins-Add-In-Katalog  **Office-Add-Ins** aus.
    

## Zusätzliche Ressourcen


- [Einrichten eines Add-In-Katalogs auf SharePoint Online](http://msdn.microsoft.com/de-de/library/1d50a571-6e02-4bc0-a3d6-6ef1eca3c2ce%28Office.15%29.aspx)
    
