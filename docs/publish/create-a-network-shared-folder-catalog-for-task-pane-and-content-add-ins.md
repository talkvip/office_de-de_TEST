
# Erstellen eines Katalogs mit freigegebenen Ordnern für Aufgabenbereich- und Inhalts-Add-ins
Richten Sie einen Katalog mit freigegebenen Ordnern für die Veröffentlichung von Aufgabenbereich- oder Inhalts-Office-Add-Ins ein.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | PowerPoint | Project | Word_

Ein Katalog mit freigegebenen Ordnern ist eine Möglichkeit, die Manifestdateien für Aufgabenbereich- und Inhalts-Office-Add-Ins in einer Netzwerkdateifreigabe zu veröffentlichen. Anschließend können Benutzer Add-Ins abrufen, indem sie mit den Schritten im folgenden Verfahren die Dateifreigabe als vertrauenswürdigen Katalog angeben.

Die Manifestdatei ist eine XML-Datei, mit der Sie deklarativ beschreiben können, wie Ihr Add-in aktiviert werden sollte, wenn es von einem Endbenutzer mit Office-Dokumenten und -Anwendungen installiert und verwendet wird. Weitere Informationen finden Sie unter [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md).

Die Manifestdatei ist die einzige Datei, die Sie im Katalog mit freigegebenen Ordnern bereitstellen sollten. Stellen Sie die Webanwendung selbst auf einem Webserver bereit, und geben Sie dessen URL in der Manifestdatei im  **SourceLocation**-Element an.

 >**Wichtig**  Um Add-ins, die auf externe Daten und Dienste zugreifen, sicherer zu machen, sollte Ihr Add-in beim Verbinden mit externen Daten und Diensten ein sicheres Protokoll verwenden, wie z. B. das Hypertext Transfer Protocol Secure (HTTPS).


## Angeben einer Dateifreigabe als vertrauenswürdigen Katalog


1. Erstellen Sie einen Ordner in einer Netzwerkfreigabe, beispielsweise  `\\MyShare\MyManifests`.
    
2. Legen Sie die zu veröffentlichenden Manifestdateien für die Aufgabenbereich- und Inhalts-Add-ins in dieser Dateifreigabe ab.
    
3. Öffnen Sie ein neues Dokument in Excel oder Word.
    
4. Wählen Sie die Registerkarte  **Datei** aus, und klicken Sie auf **Optionen**.
    
5. Klicken Sie auf  **Sicherheitscenter** und anschließend auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
    
6. Klicken Sie auf  **Vertrauenswürdige Add-in-Kataloge**.
    
7. Geben Sie im Feld  **Katalog-URL** den Pfad zu der in Schritt 1 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.
    
8. Aktivieren Sie das Kontrollkästchen  **In Menü anzeigen**, und klicken Sie auf  **OK**.
    
Nach Abschluss dieser Schritte können Sie auf der Registerkarte  **Einfügen** des Menübands die Option **Meine Add-Ins** auswählen und anschließend oben im Dialogfeld **Office-Add-Ins** die Option **Freigegebener Ordner** auswählen, um aus diesem Katalog ein Aufgabenbereich- oder Inhalts-Add-In einzufügen.

Alle weiteren Manifestdateien, die Sie in diese Dateifreigabe einfügen, sind für Benutzer verfügbar, die diesen Katalog mit freigegebenen Ordnern angegeben haben.


## Zusätzliche Ressourcen



- [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)
    
