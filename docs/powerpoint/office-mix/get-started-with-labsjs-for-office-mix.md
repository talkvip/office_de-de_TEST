
# Erste Schritte mit LabsJS für Office Mix
Erfahren Sie, wie Sie eine Office-Add-Ins-Entwicklungsumgebung einrichten können, um die Entwicklung von Microsoft Office Mix Laboren mit der Labs.js-API zu unterstützen.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

LabsJS-Inhalte machen eine API (labs.js), Beispiele, Dokumentationen und dazugehörige Dateien verfügbar, mit denen Sie interaktive Labore entwickeln können, sie in Office Mix integrieren, und anschließend in Microsoft PowerPoint rendern können. Diese Labore sind Office-Add-Ins, die Sie mit HTML5 und der labs.js JavaScript-Bibliothek erstellen können.

## LabsJS-Inhalte

LabsJS bietet Dokumentationen, Beispiel-Labore und die erforderlichen Dateien zum Erstellen und Veröffentlichen Ihrer eigenen Labore für Office Mix.


**Erforderliche Dateien**


|**Datei**|**Beschreibung**|
|:-----|:-----|
|labs-1.0.4.js|Die LabsJS JavaScript-API für die Entwicklung von Office Mix Laboren Diese Datei muss in Ihrem Projekt enthalten sein, damit es mit Office Mix integriert werden kann. Die Datei wird außerdem in dem Netzwerk für die Inhaltsübermittlung unter  `https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js` gehostet. Wenn Sie Ihre App veröffentlichen, müssen Sie eine Verknüpfung zu der Datei im CDN erstellen.|
|labs-1.0.4.d.ts|TypeScript-Defitionsdatei für labs.js Damit können Sie Ihren TypeScript-Code einfach mit labs.js integrieren. Die Definitionsdatei bietet außerdem einen umfassenden Überblick über alle in labs.js enthaltenen Komponenten. Sie können TypeScript von [http://www.typescriptlang.org/](http://www.typescriptlang.org/) herunterladen. Die Definitionsdatei wurde mit TypeScript Version 0.9.1.1 erstellt.|
|Verlauf|Versionsverlauf für die verschiedenen Versionen der labs.js-Bibliothek|
|Labshost.html|Eine Webseite, mit der Sie Ihr Labor mit Office Mix anzeigen und debuggen können, außerhalb des Kontexts von PowerPoint. Geben Sie zum Benutzern der Seite Ihre URL in das Haupteingabefenster ein und sie wird im Frame geladen. Zwischen der API und Office Mix ausgetauschte Daten, wenn es in PowerPoint oder Office Mix Lesson Player ausgeführt wird, werden in den Eingabefeldern rechts angezeigt. Die Daten können auch vorgeseedet werden. Beachten Sie, dass die Beispiel-Labore im Beispielabschnitt vorhandene Office Mix-Add-ins anzeigt, die im Hostkontext ausgeführt werden.|
|SampleManifest.xml|Ein Office-Add-Ins-Beispielmanifest, das als Vorlage zum Erstellen Ihres eigenen Anwendungsmanifests verwendet werden kann.|
|Simplelab.html|Ein mit labs.js erstelltes Beispiel-Labor Damit kann eine Webseite und die Einfügung einer Webseite ausgewählt werden, anschließend wird der Benutzer nachverfolgt, der sie anzeigt.|
|Simplelab.ts|Die zum Erstellen des simplelab-Beispiels verwendete TypeScript-Datei|
|Simplelab.js|JavaScript-Version des Simplelab-Beispiels Sie und die simplelab.ts zeigen die Nutzung der LabsJS-API.|

## Einrichten der Entwicklungsumgebung

Die labs.js-Bibliothek dient als Abstraktionsebene über der office.js-Bibliothek (die API für Office-Add-Ins), die Labore, die Sie mit der labs.js-Bibliothek erstellen, sind also tatsächlich Office-Add-Ins. Um mit der labs.js-Bibliothek arbeiten zu können und diese Labore in Office Mix auszuführen, müssen Sie sich selbst zunächst als Office-Add-Ins-Entwickler einrichten.


### Registrieren für eine Office 365-Entwicklerwebsite

Zunächst müssen Sie sich für eine Website für Office 365-Entwickler anmelden. Dadurch können Sie Ihr Labor hosten und testen, bevor Sie es an den Office Store übermitteln. Mit der Website können Sie Ihr Add-In in Office Mix veröffentlichen und in einer Live-Umgebung testen.

Weitere Informationen dazu finden Sie unter [Einrichten einer Entwicklungsumgebung für SharePoint-Add-Ins in Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). Sie müssen nur die ersten beiden Schritte befolgen; das Installieren der „Napa"-Entwicklertools ist optional.


### Einrichten eines App-Katalogs in SharePoint Online

Nachdem Ihre Entwicklerwebsite erstellt und bereitgestellt ist, müssen Sie einen Add-In-Katalog in SharePoint Online einrichten. Weitere Informationen finden Sie unter [Einrichten eines Add-In-Katalogs in Office 365](../../publish/set-up-an-add-in-catalog-on-office-365.md).

Für Office Mix verwenden Sie einen Add-In-Katalog, damit Sie die Präproduktions-Add-Ins in eine Lektion einfügen und ein umfassendes Testing durchführen, bevor Sie die Labore an den Store übermitteln.


## Erstellen des Labors

Um Ihr erstes Labor zu erstellen, befolgen Sie die Schritte in der [exemplarischen Vorgehensweise](walkthrough:-creating-your-first-lab-for-office-mix.md), in der erläutert wird, wie Sie ein einfaches wahr/falsch-Quiz erstellen können. Weitere Informationen finden Sie unter [Exemplarische Vorgehensweise: Erstellen der ersten Übungseinheit für Office Mix](walkthrough:-creating-your-first-lab-for-office-mix.md).


## Veröffentlichen des Labors

Nachdem Sie Ihr Labor erstellt haben, können Sie es veröffentlichen und an den Store übermitteln.


### Erstellen und Hochladen des Anwendungsmanifests

Das Anwendungsmanifest ist ein XML-Dokument, in dem Ihr LabJS-Labor erläutert wird. Es bietet eine Referenz zu einer URL, auf der das Labor gehostet wird und bietet Details zum Labor einschließlich des Anzeigenamens, der Beschreibung, Symbole, Größe usw.

Wir beziehen ein Beispiel-Manifest mit ein, „SampleManifest.xml". Weitere Informationen zum Manifestschema sowie einen Link zur Schemadefinition finden Sie unter [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md).

Um Ihr Manifest auf Ihre SharePoint-Website hochzuladen, müssen Sie zunächst zu Ihrem Anwendungskatalog navigieren, den Sie in der Regel unter der URL  `https://<your site>/sites/AppCatalog` finden. Klicken Sie dann auf die Schaltfläche **Neue App** und befolgen Sie die Schritte, um Ihr Anwendungsmanifest hochzuladen.


### Aktualisieren des PowerPoint 2013-Katalogs

Aktualisieren Sie als nächstes Ihren PowerPoint 2013-Katalog, nachdem Sie sich bei Ihrem Entwicklerkonto angemeldet haben.

Beginnen Sie, indem Sie Ihren PowerPoint 2013-Katalog aktualisieren. Starten Sie PowerPoint 2013 und navigieren Sie zu dem Menüpfad  **Datei > Optionen > Sicherheitscenter > Einstellungen für das Sicherheitscenter > Vertrauenswürdige App-Kataloge**. Fügen Sie nun eine Referenz zu Ihrem App-Katalog hinzu und klicken Sie auf  **OK**. PowerPoint 2013 fordert Sie auf, sich abzumelden, damit die Änderungen vorgenommen werden können.

Melden Sie sich als Letztes wieder mit Ihrem Entwicklerkonto an. Klicken Sie auf Ihren Anmeldenamen oben rechts in der Ecke in PowerPoint 2013 und melden Sie sich mit Ihrem Entwicklerkonto an. Sie können nun Ihr Add-In einfügen.


### Einfügen, Veröffentlichen und Anzeigen der App

Um Ihr Add-In in den Katalog einzufügen, klicken Sie auf das Menüband  **Einfügen** und dann auf **Store** im Abschnitt **Apps**. Klicken Sie auf  **Meine Organisation**, nun wird Ihr Add-In in Ihrem Add-In-Katalog angezeigt. Klicken Sie auf Add-In,  **Einfügen** und Ihr Add-In (Labor) wird in das PowerPoint 2013-Dokument eingefügt.

Sie können alle verfügbaren Office Mix-Funktionen nutzen, um Ihre Lektion mit Ihrem neuen Labor zu veröffentlichen.


 >**Wichtig**  Um die Anwendung anzuzeigen, müssen Sie sich bei Ihrem SharePoint-Katalog mit demselben Browser anmelden, indem Sie die Lektion anzeigen. SharePoint-Kataloge bieten nur authentifizierten Benutzern Zugriff, damit Ihnen die Anwendung angezeigt wird, bei der Sie sich zunächst anmelden müssen. 


### Übermitteln des Labors an den Office Store

Weitere Informationen zum Übermitteln Ihres Labors an den Office Store finden Sie unter [Veröffentlichen Ihres Office-Add-ins](../publish/publish.md)


## Zusätzliche Ressourcen



- [Office Mix-Add-Ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Office-Add-Ins](../../docs/overview/office-add-ins.md)
    
- [Exemplarische Vorgehensweise: Erstellen der ersten Übungseinheit für Office Mix](walkthrough:-creating-your-first-lab-for-office-mix.md)
    
