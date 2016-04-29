
# Entwicklungslebenszyklus von Office-Add-Ins
Planen des Ende-zu-Ende-Prozesses für das Entwickeln einer Aufgabenbereich-, Inhalts- und Outlook-Add-Ins zur Erweiterung von Office-Anwendungen

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Einige hilfreiche Tools zum Entwickeln eines Office-Add-Ins stellen [Visual Studio und die Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs) oder [Napa Office 365 Development Tools](https://www.napacloudapp.com/getting-Started.aspx) dar. Weitere Informationen finden Sie unter [Voraussetzungen zum Ausführen von Office-Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md). 

Der typische Entwicklungslebenszyklus eines Office-Add-Ins umfasst folgende Schritte:


1.  **Bestimmen Sie den Zweck des Add-Ins.**
    
    Stellen Sie sich die folgenden Fragen:
    
      - Welchen Nutzen hat das Add-In? 
    
  - Wie verhilft sie Ihren Kunden zu mehr Produktivität?
    
  - Welche Szenarien unterstützen die Features Ihres Add-Ins?
    

    Ermitteln Sie die wichtigsten Features und Szenarien, und konzentrieren Sie Ihren Entwurf darauf. 
    
2.  **Ermitteln Sie die Daten und die Datenquelle für das Add-In.**
    
    Befinden Sie die Daten in einem Dokument, einer Arbeitsmappe, einer Präsentation oder einer browserbasierten Access-Datenbank oder in einem oder mehreren Elementen in einem Exchange Server- oder Exchange Online-Postfach? Stammen die Daten aus einer externen Quelle wie einem Webdienst?
    
3.  **Ermitteln Sie den Typ des Add-Ins und die Office-Hostanwendung, die den Zweck des Add-Ins am besten unterstützen.**
    
    Überlegen Sie Folgendes, um die Szenarien zu ermitteln:
    
      - Verwenden Kunden das Add-In, um die Inhalte eines Dokuments oder einer browserbasierten Access-Datenbank zu erweitern? Falls ja, erwägen Sie die Erstellung eines Inhalts-Add-Ins. Derzeit können Sie Inhalts-Add-Ins für Access, Excel, Excel Online, PowerPoint oder PowerPoint Online erstellen.
    
  - Verwenden Kunden das Add-In während des Anzeigens oder Erstellens einer E-Mail oder eines Termins? Ist es wichtig, das Add-In entsprechend dem aktuellen Kontext verfügbar zu machen? Ist es eine Priorität, das Add-In nicht nur auf dem Desktop, sondern auch auf Tablets und Smartphones verfügbar zu machen?
    
    Wenn die Antwort auf eine dieser Fragen „Ja" lautet, sollten Sie die Erstellung eines Outlook-Add-Ins erwägen. Derzeit können Sie Outlook-Add-Ins für die Outlook Rich Clients, Outlook Web App und OWA für mobile Geräte erstellen, wenn sich Ihr Postfach auf einem Exchange Server befindet. Anschließend sollten Sie den Kontext ermitteln, der Ihr Add-In auslöst (beispielsweise in dem Fall, wenn der Benutzer eine Entwurfsvorlage verwendet, bestimmte Nachrichtentypen, das Vorhandensein einer Anlage, einer Adresse, eines Vorgangsvorschlags, eines Besprechungsvorschlags oder bestimmter Zeichenfolgenmuster im Inhalt einer E-Mail oder eines Termins). Weitere Informationen dazu, wie Sie das Mail-Add-In kontextspezifisch aktivieren können, finden Sie unter [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md).
    
  - Verwenden Kunden das Add-In, um das Anzeigen oder Verfassen eines Dokuments zu verbessern? Falls ja, sollten Sie das Erstellen eines Aufgabenbereich-Add-Ins in Betracht ziehen. Derzeit können Sie Aufgabenbereich-Add-Ins für Excel, Excel Online, PowerPoint, PowerPoint Online, Project und Word erstellen.
    
4.  **Entwerfen und implementieren Sie die Benutzererfahrung und die Benutzeroberfläche für das Add-In.**
    
    Entwerfen Sie eine schnelle und flüssige Benutzererfahrung, die konsistent und leicht zu erlernen ist, und für deren primäre Szenarien nur wenige Schritte erforderlich sind. Verwenden Sie je nach Zweck des Add-Ins APIs oder Webdienste von Drittanbietern.
    
    Sie können zur Implementierung der Benutzeroberfläche zahlreiche Webentwicklungstools, HTML und JavaScript verwenden.
    
5.  **Erstellen Sie basierend auf dem Office-Add-Ins-Manifestschema eine XML-Manifestdatei.**
    
    Erstellen Sie eine XML-Manifestdatei, um das Add-In und seine Anforderungen anzugeben, legen Sie die Speicherorte der vom Add-In verwendeten HTML-, JavaScript- und CSS-Dateien sowie je Add-In-Typ die Standardgröße und die Berechtigungen fest.
    
    Für Outlook-Add-Ins können Sie basierend auf der aktuellen Nachricht oder dem aktuellen Termin den Kontext angeben, unter dem Ihr Add-In relevant ist, und unter dem Outlook dieses zur Verwendung auf der Benutzeroberfläche verfügbar machen soll. Sie können auch die Geräte festlegen, die vom Add-In unterstützt werden sollen. Legen Sie in der Manifestdatei den Kontext als Aktivierungsregeln und die unterstützen Geräte fest.
    
6.  **Installieren und testen Sie das Add-In.**
    
    Legen Sie die HTML- und alle JavaScript- und CSS-Dateien auf den in der Manifestdatei des Add-Ins angegebenen Webservern ab. Der Vorgang zum Installieren eines Add-Ins ist vom Typ des Add-Ins abhängig.
    
    Installieren Sie das Outlook-Add-In in einem Exchange-Postfach, und geben Sie im Exchange Admin Center (EAC) den Speicherort der Manifestdatei des Add-Ins an. Weitere Informationen finden Sie unter [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md).
    
7.  **Veröffentlichen Sie das Add-in.**
    
    Sie können das Add-In an den Office Store übermitteln, von wo aus Kunden das Add-In installieren können. Des Weiteren können Sie Add-Ins für den Aufgabenbereich und Inhalts-Add-ins in einem privaten Add-in-Katalog-Ordner in SharePoint oder in einem freigegebenen Netzwerkordner veröffentlichen oder ein Outlook-Add-In direkt auf einem für Ihre Organisation bereitstellen. Einzelheiten hierzu finden Sie unter [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md).
    
8.  **Aktualisieren des Add-ins**
    
    Wenn von Ihrem Add-In Aufrufe an einen Webdienst durchgeführt werden und Sie nach dem Veröffentlichen des Add-Ins Aktualisierungen an dem Webdienst vornehmen, müssen Sie das Add-In nicht erneut veröffentlichen. Nehmen Sie jedoch Änderungen an jeglichen für Ihr Add-In übermittelten Elementen oder Daten wie etwa dem Add-In-Manifest, Screenshots, Symbolen, HTML- oder JavaScript-Dateien vor, müssen Sie das Add-In erneut veröffentlichen. Wenn Sie das Add-In im Office Store veröffentlicht haben, müssen Sie sie erneut übermitteln, damit die Änderungen im Office Store implementiert werden können. Sie müssen das Add-In mit einem aktualisierten Add-In-Manifest mit einer neuen Versionsnummer erneut übermitteln. Achten Sie darauf, die Versionsnummer des Add-Ins im Übermittlungsformular zu aktualisieren, sodass diese der Versionsnummer des neuen Manifests entspricht. Bei Outlook-Add-In sollten Sie sicherstellen, dass das [ID](http://msdn.microsoft.com/de-de/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)-Element eine andere UUID im Add-In-Manifest enthält.
    
