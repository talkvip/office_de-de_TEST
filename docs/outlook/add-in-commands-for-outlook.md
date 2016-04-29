
# Add-In-Befehle für Outlook
Verwenden Sie Add-In-Befehle, um Outlook-Add-Ins in die Outlook-Benutzeroberfläche zu integrieren. 

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Mit Outlook-Add-In-Befehlen können bestimmte Add-In-Aktionen im Menüband durch Hinzufügen von Schaltflächen oder Dropdownmenüs initiiert werden. Hierdurch können Benutzer auf eine einfache, intuitive und unaufdringliche Weise auf Add-Ins zugreifen. Da sie zusätzliche Funktionalität bieten, können Sie mithilfe von Add-In-Befehlen auf einfache Weise attraktivere Lösungen erstellen.

Add-In-Befehle sind nur für Add-Ins verfügbar, die nicht die Regel [ItemHasAttachment](https://msdn.microsoft.com/en-us/library/fp123567.aspx%28Office.15%29.aspx), [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/fp161166.aspx%28Office.15%29.aspx) oder [ItemHasRegularExpressionMatch](https://msdn.microsoft.com/en-us/library/fp142215.aspx%28Office.15%29.aspx) verwenden, um die Typen der Elemente zu beschränken, mit denen sie aktiviert werden. Add-ins können jedoch unterschiedliche Befehle zur Verfügung stellen, je nachdem, ob das aktuell ausgewählte Element eine Nachricht oder ein Termin ist. Außerdem können Sie in Szenarien zum Lesen oder zum Verfassen angezeigt werden. Das Verwenden von Add-In-Befehlen, wann immer dies möglich ist, ist eine [bewährte Methode](../../docs/design/add-in-development-best-practices.md).


## Erstellen eines Add-In-Befehls

Add-In-Befehle werden im Add-In-Manifest im  **VersionOverrides**-Element deklariert. Dieses Element ist eine Ergänzung zum Manifestschema v1.1, das Abwärtskompatibilität gewährleistet. In einer Umgebung, die keine  **VersionOverrides** unterstützt, werden vorhandene Add-Ins weiterhin funktionieren wie sie auch ohne Add-In-Befehle funktionieren würden.

Die  **VersionOverrides**-Manifesteinträge geben viele Aspekte des Add-In-Befehls an, z. B. den Host, die Typen von Steuerelementen, die dem Menüband hinzugefügt werden sollen, Text Symbole und alle zugehörigen Funktionen sowie den Ort, an dem der Add-In-Befehl angezeigt wird. Weitere Informationen finden Sie unter [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md). 

Wenn ein Add-In Fortschrittsaktualisierungen bereitstellen muss, z. B. Fortschrittsanzeigen oder Fehlermeldungen, muss dies über die [Benachrichtigungs-APIs](http://dev.outlook.com/reference/add-ins/NotificationMessages.html%28Office.15%29.md). Die Verarbeitung für Benachrichtigungen muss auch in einer separaten HTML-Datei definiert werden, die im  **FunctionFile**-Knoten des Manifests angegeben ist.

Entwickler sollten Symbole für alle erforderlichen Größen definieren, damit die Add-In-Befehle auf dem Menüband problemlos angepasst werden können. Symbolgrößen sind 80 x 80 Pixel, 32 x 32 Pixel und 16 x 16 Pixel.


## Wie werden die Add-In-Befehle angezeigt?

Ein Add-In-Befehl wird als Schaltfläche im Menüband angezeigt. Wenn ein Benutzer ein Add-In installiert, werden seine Befehle auf der Benutzeroberfläche als eine Gruppe von Schaltflächen angezeigt, die mit dem Add-In-Namen gekennzeichnet ist. Dies kann entweder die Standardregisterkarte oder eine benutzerdefinierte Registerkarte im Menüband sein. Für Nachrichten ist die Standardeinstellung entweder die Registerkarte  **Start** oder **Nachricht**. Für Kalender ist die Standardeinstellung die Registerkarte  **Besprechung**,  **Besprechungsereignis**,  **Besprechungsreihe** oder **Termin**. Auf der Standardregisterkarte kann jedes Add-In eine Menübandgruppe mit bis zu sechs Befehlen enthalten. Auf benutzerdefinierten Registerkarten kann das Add-In bis zu zehn Gruppen aufweisen, von denen jede sechs Befehle aufweist. Add-Ins sind auf auf benutzerdefinierte Registerkarte beschränkt.

Wird das Menüband voller, passen sich die Add-In-Befehle an (ausblenden). Die Add-In-Befehle eines Add-Ins werden immer zu einer Gruppe zusammengefasst.


![Screenshot mit den normalen und den reduzierten Add-In-Befehlsschaltflächen.](../../images/6fcb64d8-9598-41d1-8944-f6d1f6d2edb6.png)Nachdem ein Add-In Befehl zu einem Add-In hinzugefügt wurde, wird der Name des Add-Ins aus der App-Leiste entfernt, es sei denn, das Add-In enthält auch einen [Benutzerdefinierter Bereich Outlook-Add-Ins](../outlook/custom-pane-outlook-add-ins.md). Nur die Add-In-Befehlsschaltfläche auf dem Menüband bleibt bestehen.


## Welche UX-Shapes sind für Add-In-Befehle vorhanden?

Das UX-Shape für ein Add-In-Befehl besteht aus einer Registerkarte des Menübands in der Hostanwendung, die die Schaltflächen enthält, mit denen verschiedene Funktionen ausgeführt werden können. Derzeit werden drei Benutzeroberflächenformen unterstützt:


- Eine Schaltfläche, mit der eine JavaScript-Funktion ausgeführt wird
    
- Eine Schaltfläche, mit der ein Aufgabenbereich gestartet wird
    
- Eine Schaltfläche, die ein Dropdownmenü mit einer oder mehrerer Schaltflächen der anderen beiden Typen anzeigt
    

### Ausführen einer JavaScript-Funktion

Verwenden Sie eine Add-In-Befehlsschaltfläche, mit der eine JavaScript-Funktion ausgeführt wird, für Szenarien, in denen Benutzer zur Initiierung keine zusätzliche Auswahl treffen müssen. Dies gilt z. B. für Aktionen wie Nachverfolgen, Erinnern oder Drucken oder wenn Benutzer ausführliche Aktionen zu einem Dienst möchte. 


![Eine Schaltfläche, die eine Funktion im Outlook-Menüband ausführt.](../../images/23ab1de3-3ec4-41a5-ba5b-30b11d464e0c.png)


### Starten eines Aufgabenbereichs

Verwenden Sie eine Add-In-Befehlsschaltfläche zum Starten eines Aufgabenbereichs für Szenarien, in denen ein Benutzer mit einem Add-In über einen längeren Zeitraum interagieren muss. Beispiel: Für das Add-In sind möglicherweise Änderungen an Einstellungen oder das Ausfüllen vieler Felder erforderlich. 

Die Standardbreite des vertikalen Aufgabenbereichs beträgt 300 px. Die Größe des vertikalen Aufgabenbereichs kann im Outlook Explorer und Inspektor geändert werden. Der Bereich kann auf dieselbe Weise die Größe von Aufgabenbereich und Listenansicht ändern.


![Eine Schaltfläche, die einen Aufgabenbereich im Outlook-Menüband öffnet.](../../images/c8e03da8-9f71-4f9b-813f-1cdea43d433c.png)Der obige Screenshot zeigt ein Beispiel für den vertikalen Aufgabenbereich. Der Bereich wird mit dem Namen des Add-In-Befehls in der oberen linken Ecke angezeigt. Es gibt auch eine  **X**-Schaltfläche in der oberen rechten Ecke des Bereichs, damit der Benutzer das Add-In schließen kann, wenn er es nicht mehr verwendet. Dieser Bereich wird nicht über alle Nachrichten hinweg beibehalten. Jedes Rendern von Benutzeroberflächenelementen im Aufgabenbereich, abgesehen vom Namen und der Schaltfläche zum Schließen, erfolgt durch das Add-In. 

Wenn ein Benutzer auf einen anderen Add-In-Befehl klickt, mit dem ein Aufgabenbereich geöffnet wird, wird der Aufgabenbereich durch den Befehl ersetzt, auf den zuletzt geklickt wurde. Wenn ein Benutzer auf eine Add-In-Befehlsschaltfläche, mit der eine Funktion ausgeführt wird, oder auf ein Dropdownmenü klickt, wenn der Aufgabenbereich geöffnet ist, wird die Aktion fertig gestelt und der Aufgabenbereich bleibt geöffnet.


### Dropdownmenü

Ein Add-In-Befehl für ein Dropdownmenü definiert eine statische Liste von Schaltflächen, die innerhalb des Dropdownmenüs dargestellt werden. Die Schaltflächen in dem Menü können eine Kombination von Schaltflächen sein, die eine Funktion ausführen, oder Schaltflächen, die einen Aufgabenbereich öffnen. Untermenüs werden nicht unterstützt.


![Eine Schaltfläche, die ein Dropdown-Menü im Outlook-Menüband öffnet.](../../images/3eff90d6-7822-4fdb-9153-68f754c0c746.png)


## Wo werden die Add-In-Befehle auf der Benutzeroberfläche angezeigt?

Add-In-Befehle werden für vier Szenarien unterstützt:


### Lesen einer Nachricht

Wenn der Benutzer eine Nachricht liest, werden Add-In-Befehle, die zur Standardregisterkarte hinzugefügt werden, auf der Registerkarte  **Start** angezeigt, wenn die Nachricht im Lesebereich angezeigt wird, und auf der Registerkarte **Nachricht** für ein Popout-Leseformular.


### Verfassen einer Nachricht

Wenn der Benutzer eine Nachricht verfasst, werden der Standardregisterkarte hinzugefügte Add-In-Befehle auf der Registerkarte  **Nachricht** angezeigt.


### Erstellen oder Anzeigen eines Termins oder einer Besprechung als Organisator

Beim Erstellen oder Anzeigen eines Termins oder einer Besprechung als Organisator, werden der Standardregisterkarte hinzugefügte Add-In-Befehle auf den Registerkarten  **Besprechung**,  **Besprechungsserienelement**,  **Besprechungsserie** oder **Termine** auf Popoutformularen angezeigt. Wenn der Benutzer jedoch ein Element im Kalender auswählt, das Popout aber nicht öffnet, wird die Menübandgruppe des Add-Ins nicht im Menüband angezeigt.


### Anzeigen einer Besprechung als Teilnehmer

Beim Anzeigen einer Besprechung als Teilnehmer werden der Standardregisterkarte hinzugefügte Add-In-Befehle auf den Registerkarten  **Besprechung**,  **Besprechungsserienelement**,  **Besprechungsserie** auf Popoutformularen angezeigt. Wenn der Benutzer jedoch ein Element im Kalender auswählt, das Popout aber nicht öffnet, wird die Menübandgruppe des Add-Ins nicht im Menüband angezeigt.


## Zusätzliche Ressourcen



- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md)
    
- [Add-in Command Demo Outlook Add-in](https://github.com/jasonjoh/command-demo.aspx)
    
