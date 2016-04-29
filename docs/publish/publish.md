
# Veröffentlichen Ihres Office-Add-ins
Machen Sie Ihre Add-Ins über einen Add-in-Katalog, einen freigegebenen Ordner oder einen Exchange-Server für Benutzer verfügbar. Erweitern Sie die Reichweite Ihrer Add-Ins für die Endbenutzer.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Sie können Ihre Add-Ins im Office Store veröffentlichen oder in einen privaten freigegebenen Ordnerkatalog für Add-ins in SharePoint, in einen freigegebenen Netzwerkordner oder auf einen Exchange-Server hochladen. Die verfügbaren Optionen sind vom Typ des erstellten Add-Ins abhängig.

Weitere Informationen zum Veröffentlichen im Office Store finden Sie unter [Submit add-ins and web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).


**Optionen für die Veröffentlichung von Office-Add-Ins**


|**Typ**|**Office Store**|**Unternehmens-Add-In-Katalog**|**Add-In-Katalog mit freigegebenen Ordnern**|**Exchange-Server**|
|:-----|:-----|:-----|:-----|:-----|
|Aufgabenbereich-Add-In|x|x|x||
|Inhalts-Add-In|x|x|x||
|Outlook Add-In|x|||x|
Vor der Veröffentlichung Ihrer Add-In müssen Sie sie [verpacken](../publish/package-your-add-in-using-napa-or-visual-studio.md). Außer der Veröffentlichung Ihrer Add-Ins für Endbenutzer sollten Sie auch überlegen, wie Sie die Reichweite der Add-Ins vergrößern können.


## Veröffentlichen von Aufgabenbereich- und Inhalts-Add-ins in einem Add-in-Katalog


Für Aufgabenbereich- und Inhalts-Add-ins können IT-Abteilungen private Unternehmens-Add-in-Kataloge bereitstellen und konfigurieren, um dieselbe Office-Lösungskatalogerfahrung wie mit dem Office Store zu ermöglichen. Mithilfe dieser neuen Katalog- und Entwicklungsplattform kann die IT-Abteilung eine optimierte Methode verwenden, um Office- und SharePoint-Add-Ins für verwaltete Benutzer zentral bereitzustellen, ohne dass auf jedem Client Lösungen installiert werden müssen. Anschließend können Sie mit dem Telemetrietool die Add-in-Verwendung überwachen, die Kompatibilität überprüfen und Endbenutzerprobleme behandeln. Weitere Informationen finden Sie unter:


- [Veröffentlichen von Aufgabenbereich- und Inhalts-Add-Ins in einem Add-In-Katalog in SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)
    
- [Einrichten eines Add-In-Katalogs unter SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md)
    
- [Einrichten eines Add-In-Katalogs in Office 365](../../publish/set-up-an-add-in-catalog-on-office-365.md)
    

## Veröffentlichen von Aufgabenbereich- und Inhalts-Add-ins in einem freigegebenen Netzwerkordner


In einem Unternehmensumfeld kann die IT-Abteilung alternativ Aufgabenbereich- und Inhalts-Add-ins, die von internen oder externen Entwicklern erstellt wurden, in einem freigegebenen Netzwerkordner bereitstellen, in dem die Manifestdateien gespeichert und verwaltet werden. Wenn Entwickler ihre Add-ins später aktualisieren, müssen in beiden Fällen keine Updates an die Endbenutzer übermittelt werden bzw. muss die IT-Abteilung die Add-ins nicht erneut für die Unternehmensbenutzer bereitstellen. Informationen zum Einrichten eines Add-in-Katalogs in einem freigegebenen Netzwerkordner finden Sie unter [Erstellen eines Katalogs mit freigegebenen Ordnern für Aufgabenbereich- und Inhalts-Add-ins](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).


## Veröffentlichen von Outlook-Add-Ins auf einem Exchange-Server


Outlook-Add-Ins werden in einem Exchange-Katalog veröffentlicht, der für Benutzer des Exchange-Servers verfügbar ist, auf dem er gespeichert ist. Dies ermöglicht das Veröffentlichen und Verwalten von Outlook-Add-Ins. Hierzu gehören intern erstellte Outlook-Add-Ins sowie Lösungen, die vom Office Store bezogen und für die Verwendung im Unternehmen lizenziert sind. Mail-Add-ins werden mit dem Exchange Admin Center (EAC) oder durch Ausführen von Windows PowerShell-Remotebefehlen (Cmdlets) in einem Exchange-Katalog installiert. Informationen zum Veröffentlichen eines Outlook-Add-Ins finden Sie unter [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md).


## Add-in-Benutzererfahrung für die Endbenutzer


Für Endbenutzer ist es einfach, Add-ins zu erwerben, zu installieren und auszuführen. Die Benutzererfahrung ist gleich, unabhängig davon, wie der Zugriff auf die Add-ins erfolgt:


- Über den öffentlichen Office Store, mithilfe ihres Microsoft-Kontos
    
- Über einen SharePoint-Add-in-Katalog, mithilfe Ihrer Unternehmens-ID
    
- Über einen freigegebenen Netzwerkordner
    
- Über einen Exchange-Server
    
Beispielsweise melden sich die Benutzer für den Erwerb eines neuen Aufgabenbereich-Add-Ins in Excel bei Office mit ihrem Microsoft-Konto an, öffnen eine Excel-Arbeitsmappe und wählen auf der Registerkarte  **Einfügen** des Menübands die Option **Meine Add-Ins** aus. Daraufhin wird das Dialogfeld **Office-Add-Ins** angezeigt.

Im Dialogfeld  **Office-Add-In** wählt der Benutzer **Nach weiteren Add-ins im Office Store suchen** aus. Nachdem sich die Benutzer bei Office.com mit demselben Microsoft-Konto angemeldet haben, können sie das gewünschte Add-in herunterladen und per Kreditkarte bezahlen.

In Excel wählt der Benutzer im Dialogfeld  **Office-Add-Ins** die Option **Aktualisieren**, das heruntergeladene Add-in und  **Einfügen** aus.

Nach der Anmeldung bei ihrem Konto haben Benutzer von jedem Computer und überall Zugriff auf diese Add-ins, auch auf Computern mit Office 365.


## Erweitern der Reichweite für Ihr Add-In


Um dafür zu sorgen, dass Ihre Add-In mehr Endbenutzer erreicht, sollten Sie sicherstellen, dass sie plattformübergreifend funktioniert. Die Office.js-Version 1.1 bietet Unterstützung für Office Online, und beim Validierungsprozess durch den Office Store wird die Add-In-Unterstützung für Office Online überprüft. Testen Sie Ihre Add-In vor dem Veröffentlichen, um sicherzustellen, dass sie in Office Online funktioniert.

Informationen zum Verfügbarmachen Ihres Add-Ins im Office Store finden Sie unter [Submit add-ins and web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).

Office-Add-Ins werden auch in Office für iPad unterstützt. Testen Sie Ihre Add-In unbedingt auf dem iPad, bevor Sie sie [an das Verkäuferdashboard übermitteln](http://msdn.microsoft.com/library/260ef238-0be4-42d6-ba15-1249a8e2ff12%28Office.15%29.aspx). Wenn Sie sich vergewissert haben, dass Ihre Add-In wie erwartet ausgeführt wird, können Sie Ihre Übermittlung im Verkäuferdashboard als iOS-kompatibel markieren. Zur Überprüfung müssen Sie Ihre Apple Developer-ID angeben. Weitere Informationen finden Sie unter [Debuggen von Office-Add-Ins auf dem iPad und einem Mac-Computer](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md).

Informationen zur Behebung von Benutzerproblemen mit Ihren Add-Ins finden Sie unter [Problembehandlung von Benutzerfehlern mit Office Add-Ins](../../docs/testing/testing-and-troubleshooting.md). Sie können auch [direkt auf Kundenbewertungen im Office Store reagieren](https://msdn.microsoft.com/library/jj635874.aspx).




## Weitere Ressourcen



- [Office-Add-Ins](../../docs/overview/office-add-ins.md)
    
- [Veröffentlichen von Office- und SharePoint-Add-Ins und Office 365 Web Apps im Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    


