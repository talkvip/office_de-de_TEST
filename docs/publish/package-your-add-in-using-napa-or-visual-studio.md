
# Verpacken des Add-Ins mit Napa oder Visual Studio für die Veröffentlichung
Verpacken Sie Ihr Office-Add-In für die Veröffentlichung. 

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Ihr Office-Add-In enthält eine XML-Datei, die Sie verwenden, um das Add-In zu veröffentlichen. Sie müssen die Webanwendungsdateien Ihres Projekts separat veröffentlichen.

## Verpacken einer Office-Add-In, die Sie mithilfe von Napa erstellen



1. Wählen Sie in Napa am Rand der Seite die Schaltfläche  **Veröffentlichen** (
![Schaltfläche "Veröffentlichen"](../../images/Apps_NAPA_Publish.png)).
    
2. Wählen Sie im Dialogfeld  **Veröffentlichungseinstellungen** die Schaltfläche **Weiter**.
    
3. Geben Sie die URL der Website an, die die Inhaltsdateien Ihres Add-Ins hostet (beispielsweise die default.html- und JavaScript-Dateien Ihres Projekts), und klicken Sie anschließend auf  **Veröffentlichen**.
    
4. Wählen Sie in dem Dialogfeld  **Veröffentlichung erfolgreich** den Link **Veröffentlichungsort**.
    
    Daraufhin erscheint eine Dokumentenbibliothek, die die XML-Manifestdatei Ihres Add-ins und die Webinhaltsdateien enthält. 
    
Kopieren Sie dann die Webinhaltsdateien (Stylesheerts, JavaScript-Dateien und HTML-Dateien) manuell auf den Webserver, der die von Ihnen im Dialogfeld  **Veröffentlichungseinstellungen** angegebene Website hostet.

Sie können nun Ihr XML-Manifest an den entsprechenden Speicherort hochladen, um [Ihr Add-In zu veröffentlichen](../publish/publish.md). 


## Bereitstellen des Webprojekts und Verpacken des Add-Ins mithilfe von Visual Studio 2015



### Bereitstellen Ihres Webprojekts


1. Öffnen Sie im Projektmappen-Explorer das Kontextmenü für die App für Office-Projekt und wählen Sie  **Veröffentlichen** aus.
    
    Die Seite  **Add-in veröffentlichen** wird angezeigt.
    
2. Wählen Sie in der Dropdownliste  **Aktuelles Profil** ein Profil oder wählen Sie **&lt;Neu …&gt;**, um ein neues Profil zu erstellen.
    
     >**Hinweis**  Ein Veröffentlichungsprofil gibt den Server an, auf dem Sie etwas bereitstellen und außerdem die Anmeldeinformationen für den Server, die Datenbanken für die Bereitstellung sowie weitere Bereitstellunsgoptionen.

    Wenn Sie  **&lt;Neu …&gt;** wählen, erscheint der **Veröffentlichungsprofil erstellen**-Assistent. Sie können diesen Assistenten verwenden, um ein Veröffentlichungsprofil von einem Website-Hostinganbieter wie Microsoft Azure zu importieren, oder ein neues Profil erstellen und dort Ihrem Server, Anmeldeinformationen und weitere Einstellungen in dem nächsten Prozedur hinzufügen.
    
    Weitere Informationen zum Importieren von Veröffentlichungsprofilen oder das Erstellen neuer Veröffentlichungsprofile finden Sie unter [Erstellen eines Veröffentlichungsprofils](http://msdn.microsoft.com/de-de/library/dd465337.aspx#creating_a_profile).
    
3. Wählen Sie auf der Seite  **Add-in veröffentlichen** den Link **Webprojekt bereitstellen** aus.
    
    Nun erscheint das Dialogfeld  **Web veröffentlichen**. Weitere Informationen zurt Nutzung des Assistenten finden Sie unter [Kurzanleitung: Bereitstellen eines Webprojekts mithilfe von On-Click Publishing in Visual Studio](http://msdn.microsoft.com/de-de/library/dd465337.aspx).
    

### So verpacken Sie das Add-in


1. Wählen Sie auf der Seite  **Add-in veröffentlichen** den Link **Add-in verpacken** aus.
    
    Der Assistent zum  **Veröffentlichen von Office- und SharePoint-Add-ins** wird angezeigt.
    
2. Wählen Sie in der Dropdownliste  **Wo wird Ihre Website gehostet?** die URL der Website aus, die die Inhaltsdateien Ihrer Add-In hosten wird, oder geben Sie diese ein. Wählen Sie dann **Fertig stellen**.
    
    Sie müssen eine Adresse angeben, die mit dem Präfix „HTTPS" beginnt, um den Assistenten abzuschließen. Grundsätzlich ist die Verwendung eines HTTPS-Endpunkts für Ihre Website die beste Vorgehensweise, wenn Sie jedoch nicht planen, Ihr Add-In im Office Store zu veröffentlichen, ist es nicht erforderlich. Nachdem das Paket erstellt wurde, können Sie das Manifest im Editor öffnen und das Präfix „HTTPS" Ihrer Website durch „HTTP" ersetzen. Weitere Informationen finden Sie unter [Warum müssen meine Add-Ins SSL-gesichert sein?](http://msdn.microsoft.com/de-de/library/jj591603#bk_q7). 
    
     >**Hinweis**  Azure-Websites stellen automatisch einen HTTPS-Endpunkt bereit.

    Visual Studio generiert alle Dateien, die Sie für die Veröffentlichung Ihres Add-ins benötigen, und öffnet dann den Ordner „Ausgabe veröffentlichen". 
    
Falls Sie planen, Ihr Add-In im Office Store zu veröffentlichen, können Sie den Link  **Validierungsüberprüfung ausführen** auswählen, um Probleme zu erkennen, die verhindern würden, dass Ihr Add-In akzeptiert wird. Sie sollten alle Probleme vor der Übermittlung Ihres Add-ins in den Store beheben.

Sie können nun Ihr XML-Manifest an den entsprechenden Speicherort hochladen, um [Ihr Add-In zu veröffentlichen](../publish/publish.md). Sie finden das XML-Manifest in  `OfficeAppManifests` im Ordner `app.publish`.

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## Zusätzliche Ressourcen



- [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)
    
- [Veröffentlichen von Office- und SharePoint-Add-Ins und Office 365 Web Apps im Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
