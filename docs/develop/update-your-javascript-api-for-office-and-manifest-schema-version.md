
# Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien
Aktualisieren Sie die JavaScript-API für Office-Dateien (Office.js und App-spezifische .js-Dateien) und die Add-In-Manifestvalidierungsdatei (offappmanifest-1.1.xsd) im Office-Add-In **REMOVE_ME** -Projekt von Version 1.0 auf Version 1.1.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_


## Aktuelleste Projektdateien verwenden


Wenn Sie Visual Studio zum Entwicklen Ihrer Add-In verwenden, müssen Sie, um die [neuesten API-Elemente](http://msdn.microsoft.com/de-de/library/802cf4ae-7c18-4e7d-b4d6-ecaa84c569bc%28Office.15%29.aspx) der JavaScript-API für Office und die [1.1-Features des Add-In-Manifests](4139ff24-afac-472a-af7d-9d069587ac9b.md) (das anhand von offappmanifest-1.1.xsd validiert wird) zu nutzen, [Visual Studio 2015 und die neuesten Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) herunterladen und installieren.

Wenn Sie einen Text-Editor oder eine andere IDE als Visual Studio zum Entwickeln der Add-In verwenden, müssen Sie die Verweise auf das CDN für Office.js und die Version des Schemas, auf das im Manifest der Add-In verwiesen wird, aktualisieren.

Um ein Add-in auszuführen, das mit neuen und aktualisierten Office.js-API- und App-Manifestfeatures entwickelt wurde, müssen Ihre Kunden lokale Office 2013 SP1-Produkte (oder eine höhere Version) und ggf. SharePoint Server 2013 SP1 und verwandte Serverprodukte, Exchange Server 2013 Service Pack 1 (SP1) oder gleichwertige im Internet gehostete Produkte Office 365, SharePoint Online und Exchange Online ausführen.

Informationen zum Herunterladen von Office, SharePoint und Exchange SP1-Produkten finden Sie in folgenden Ressourcen:


- [Liste aller SP1-Updates (Service Pack 1) für Microsoft Office 2013 und dazugehörige Desktopprodukte](http://support.microsoft.com/kb/2850036)
    
- [Liste aller SP1-Updates (Service Pack 1) für Microsoft SharePoint Server 2013 und dazugehörige Serverprodukte](http://support.microsoft.com/kb/2850035)
    
- [Beschreibung von Exchange Server 2013 Service Pack 1](http://support.microsoft.com/kb/2926248)
    

## Aktualisieren eines mit Visual Studio erstellten Office-Add-In-Projekts für die Verwendung der neuesten Bibliothek der JavaScript-API für Office und des Add-In-Manifestschemas der Version 1.1


Bei Projekten, die vor Version 1.1 der JavaScript-API für Office erstellt wurden, können Sie die Projektdateien mit dem  **NuGet-Paket-Manager**aktualisieren und dann die HTML-Seiten des Add-ins aktualisieren, um auf diese Dateien zu verweisen. 

Beachten Sie, dass das Update  _pro Projekt_ angewendet wird. Das heißt, Sie müssen den Updatevorgang für jedes Add-in-Projekt wiederholen, in dem Sie die Version 1.1 von Office.js und des Add-in-Manifestschemas verwenden möchten.




### So aktualisieren Sie die Bibliotheksdateien der JavaScript-API für Office in Ihrem Projekt auf die neueste Version


1. Öffnen Sie in Visual Studio 2015 das  **Office-Add-In**-Projekt, oder erstellen Sie ein neues.
    
      - Klicken Sie im linken Bereich auf  **Aktualisieren**, und schließen Sie den Paketaktualisierungsvorgang ab.
    
  - Fahren Sie mit Schritt 6 fort.
    
2. Wählen Sie die Option  **Tools** > **NuGet-Paket-Manager** > **NuGet-Pakete für Projektmappe verwalten**.
    
3. Wählen Sie im  **NuGet-Paket-Manager** **nuget.org** für **Paketquelle** und **Upgrade verfügbar** für **Filter**, und wählen Sie dann „Microsoft.Office.js".
    
4. Klicken Sie im linken Bereich auf  **Aktualisieren**, und stellen Sie dann den Paketaktualisierungsvorgang fertig.
    
5. Kommentieren Sie auf den HTML-Seiten des Add-Ins im  **&lt;head&gt;**-Tag vorhandene Skriptverweise auf office.js aus, oder löschen Sie sie (Beispiel:  `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`, und referenzieren Sie dann die aktualisierte Bibliothek der JavaScript-API für Office wie folgt (ändern Sie den Versionswert in  `1`):
    
  ```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```


    Die  `/1/` vor `office.js` in der CDN-URL gibt an, dass die letzte inkrementelle Version innerhalb der Version 1 von Office.js verwendet wird.
    

### So aktualisieren Sie die Manifestdatei in Ihrem Projekt für die Verwendung von Schema Version 1.1


- Aktualisieren Sie in der Add-In-Manifestdatei Ihres Projekts ( _Projektname_ Manifest.xml) das **xmlns**-Attribut des  **OfficeApp**-Elements, indem Sie den Versionswert in  `1.1` ändern (die Attribute außer **xmlns** bleiben unverändert):
    
  ```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
  ```


     >**Hinweis**  Nach dem Aktualisieren der Version des Add-In-Manifestschemas auf 1.1 müssen Sie die Elemente  **Capabilities** und **Capability** entfernen und durch die [Elemente Hosts und Host](cff9fbdf-a530-4f6e-91ca-81bcacd90dcd.md) oder die [Elemente Requirements und Requirement](6b6702f2-b0a5-46ab-a356-8dda897ca8ae.md) ersetzen.

## Aktualisieren eines mit einem Text-Editor oder einer anderen IDE erstellten Office-Add-In-Projekts für die Verwendung der neuesten Version der Bibliothek der JavaScript-API für Office und Version 1.1 des Add-in-Manifestschemas


Bei vor Version 1.1 der JavaScript-API für Office und des Add-in-Manifestschemas erstellten Projekten müssen Sie die HTML-Seiten Ihres Add-ins so aktualisieren, dass sie auf das CDN der v1.1-Bibliothek verweisen, und die Manifestdatei des Add-ins zur Verwendung von Schema v1.1 aktualisieren. 

Beachten Sie, dass das Update  _pro Projekt_ angewendet wird. Das heißt, Sie müssen den Updatevorgang für jedes Add-in-Projekt wiederholen, in dem Sie die Version 1.1 von Office.js und des Add-in-Manifestschemas verwenden möchten.


 >**Hinweis**  Sie benötigen keine lokalen Kopien der JavaScript-API für Office-Dateien (Office.js und App-spezifische .js-Dateien) zum Entwickeln einer Office-Add-In (bei einem Verweis auf das CDN für Office.js werden die erforderlichen Dateien zur Laufzeit heruntergeladen). Sie können jedoch auf Wunsch eine lokale Kopie der Bibliotheksdateien mit dem [NuGet-Befehlszeilenprogramm](http://docs.nuget.org/consume/installing-nuget) und dem Befehl `Install-Package Microsoft.Office.js` herunterladen.Eine Kopie der XSD (XML-Schemadefinition) für das Add-in-Manifest v1.1 finden Sie in der Auflistung unter [Schemazuordnung (App-Manifestschema v1.1)](http://msdn.microsoft.com/library/d5f72bff-3446-c64f-02ca-ab10b5648789%28Office.15%29.aspx).


### So aktualisieren Sie die Bibliotheksdateien der JavaScript-API für Office in Ihrem Projekt auf die neueste Version


1. Öffnen Sie die HTML-Seiten für Ihr Add-in in Ihrem Text-Editor oder Ihrer IDE.
    
2. Kommentieren Sie auf den HTML-Seiten des Add-Ins im  **&lt;head&gt;**-Tag vorhandene Skriptverweise auf office.js aus, oder löschen Sie sie (Beispiel:  `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`, und referenzieren Sie dann die aktualisierte Bibliothek der JavaScript-API für Office wie folgt (ändern Sie den Versionswert in  `1`):
    
  ```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```


    Die  `/1/` vor `office.js` in der CDN-URL gibt an, dass die letzte inkrementelle Version innerhalb der Version 1 von Office.js verwendet wird.
    

### So aktualisieren Sie die Manifestdatei in Ihrem Projekt für die Verwendung von Schema Version 1.1


- Aktualisieren Sie in der Add-In-Manifestdatei Ihres Projekts ( _Projektname_ Manifest.xml) das **xmlns**-Attribut des  **OfficeApp**-Elements, indem Sie den Versionswert in  `1.1` ändern (die Attribute außer **xmlns** bleiben unverändert):
    
  ```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
  ```


    Beachten Sie, dass Sie nach dem Aktualisieren der Version des Add-in-Manifestschemas auf 1.1 die  **Capabilities**- und  **Capability**-Elemente entfernen und durch die [Hosts- und Host-Elemente](cff9fbdf-a530-4f6e-91ca-81bcacd90dcd.md) oder die [Requirements- und Requirement-Elemente](6b6702f2-b0a5-46ab-a356-8dda897ca8ae.md) ersetzen müssen.
    

## Zusätzliche Ressourcen



- [Angeben von Office-Hosts und API-Anforderungen](6b6702f2-b0a5-46ab-a356-8dda897ca8ae.md)
    
- [Grundlegendes zur JavaScript-API für Office](01180dae-ca45-40c8-b3dd-fd2a85651c0c.md)
    
- [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)
    
- [Schemareferenz für Office-Add-in-Manifeste (v1. 1)](http://msdn.microsoft.com/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92%28Office.15%29.aspx)
    
