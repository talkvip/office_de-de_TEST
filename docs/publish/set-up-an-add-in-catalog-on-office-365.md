
# Einrichten eines Add-In-Katalogs in Office 365
Richten Sie einen Add-In-Katalog unter Office 365 (SharePoint Online) zum Veröffentlichen von Aufgabenbereich- und Inhalts-Office-Add-Ins ein.

 _**Gilt für:** apps for Office | Office Add-ins_

Ein Add-In-Katalog ist eine dedizierte Websitesammlung in einer SharePoint-Webanwendung (oder einem SharePoint Online-Mandanten), die als Host für Dokumentbibliotheken für SharePoint-Add-Ins und Office-Add-Ins dient. Administratoren können Office-Add-Ins-Manifestdateien in den Add-In-Katalog hochladen, um sie in ihrer Organisation zu verwenden. Wenn ein Administrator einen Add-In-Katalog als einen vertrauenswürdigen Katalog registriert (indem er Gruppenrichtlinien festlegt oder indem er den vertrauenswürdigen Katalog auf der Registerkarte  **Vertrauenswürdige Add-In-Kataloge** im Dialogfeld **Optionen** angibt, und zwar über **Datei** > **Optionen** > **Sicherheitscenter** > **Einstellungen für das Sicherheitscenter** > **Vertrauenswürdige Add-In-Kataloge**), können Benutzer das Add-In über den Einfügebereich einer Office-Clientanwendung einfügen.

 >**Hinweis**  Der Name "Apps für Office" wird in "Office-Add-Ins" geändert. Während dieses Übergangs kann in der Dokumentation und in der Benutzeroberfläche einiger Office-Anwendungen und Visual Studio-Tools möglicherweise noch der Begriff "App/Apps" verwendet werden. Weitere Informationen finden Sie unter [Neuer Name für Apps für Office und SharePoint: Office- und SharePoint-Add-Ins](https://msdn.microsoft.com/de-de/library/fp161507.aspx#Anchor_2).


## So richten Sie einen Add-In-Katalog in SharePoint Online ein


1. Wählen Sie auf der Seite mit dem Office 365 Admin Center  **Administrator** und dann **SharePoint** aus.
    
2. Wählen Sie im linken Aufgabenbereich  **Add-ins** aus.
    
3. Wählen Sie auf der Seite  **Add-ins** die Option **Add-in-Katalog** aus.
    
4. Wählen Sie auf der Seite  **Add-in-Katalogwebsite** **OK** aus, um die Standardoption zu akzeptieren und eine neue Add-in-Katalogwebsite zu erstellen.
    
5. Geben Sie auf der Seite  **Add-in-Katalog-Websitesammlung erstellen** den Titel Ihrer Add-in-Katalogwebsite an.
    
6. Geben Sie die Websiteadresse an.
    
7. Stellen Sie das  **Speicherkontingent** auf den niedrigsten Wert ein (momentan 110). Sie werden Add-In-Pakete ausschließlich auf dieser Websitesammlung installieren, und diese Pakete sind sehr klein.
    
8. Stellen Sie das  **Serverressourcenkontingent** auf 0 (null) ein. (Das Serverressourcenkontingent hat Auswirkungen auf die Einschränkung leistungsschwacher Sandkastenlösungen; Sie müssen jedoch keine Sandkastenlösungen auf der Add-in-Katalogwebsite installieren.)
    
9. Wählen Sie  **OK** aus.
    
Zum Hinzufügen von Add-In zur Add-In-Katalogwebsite navigieren Sie zur gerade erstellten Website. Wählen Sie im linken Navigationsbereich  **Office-Add-Ins** aus, und wählen Sie dann **Neues Add-In** aus, um eine Office-Add-In-Manifestdatei hochzuladen.


## Zusätzliche Ressourcen



- [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)
    
- [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md)
    
