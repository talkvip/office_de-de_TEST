
# Erstellen und Debuggen von Office-Add-Ins in Visual Studio
Verwenden Sie Visual Studio-Projektvorlagen, um ein Office-Add-Ins-Projekt zu erstellen, die Add-In-Einstellungen zu ändern und Ihr Add-In zu entwickeln und zu debuggen.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_


 >**Hinweis**  Diese Anweisungen basieren auf Visual Studio 2015. Andere Versionen von Visual Studio können von diesen Anweisungen etwas abweichen.



## Erstellen Sie ein Office-Add-In-Projekt in Visual Studio.


Beginnen Sie mit der  **Office-Add-Ins**-Projektvorlage in Visual Studio. Der Assistent zum  **Erstellen von Office-Add-In** fordert Sie auf, den Typ für das Add-in festzulegen, das Sie erstellen möchten. Anschließend stellt er Optionen für die Standardkonfiguration Ihres Add-ina bereit.


1. Wählen Sie in der Menüleiste Visual Studio **Datei**,  **Neu**,  **Projekt** aus.
    
2. Erweitern Sie in der Liste der Projekttypen unter  **Visual C#** oder **Visual Basic** die Option **Office/SharePoint**, wählen Sie  **Office-Add-Ins** aus, und wählen Sie dann **Office-Add-In** aus.
    
3. Geben Sie einen Namen für das Projekt ein, und klicken Sie dann auf  **OK**.
    
4. Das Dialogfeld  **Office-Add-in erstellen** wird geöffnet. Wählen Sie den Typ des zu erstellenden Add-ins und dann die Schaltfläche **Weiter** aus.
    
5. Wählen Sie die Standardoptionen für das zu erstellende Add-in und dann  **Fertig stellen** aus.
    
    Visual Studio erstellt das Projekt, und die Dateien werden im  **Projektmappen-Explorer** angezeigt. Die Standardseite Home.html wird in Visual Studio geöffnet.
    
In Visual Studio 2015 wurden einige der Add-In-Projektvorlagen aktualisiert und bieten Ihnen nun zusätzliche Funktionen:


- Inhalts-Add-ins können im Textkörper von Access- und PowerPoint-Dokumenten sowie Excel-Arbeitsblättern angezeigt werden. Sie können außerdem die Option „Basic Project" auswählen, um ein einfaches Inhalts-Add-in mit minimalem Startercode zu erstellen, oder Sie wählen die Option „Document Visualization Project" (nur für Access und Excel), um ein umfassenderes Inhalts-Add-in zu erstellen, das Startercode zum Visualisieren und Binden an Daten enthält.
    
- Outlook-Add-Ins umfassen Optionen, die Ihr Add-In nicht nur in E-Mail-Nachrichten und Termine einbindet, sondern auch angeben, ob das Add-In zur Verfügung steht, wenn eine E-Mail-Nachricht oder ein Termin verfasst bzw. gelesen wird.
    

 >**Hinweis**  Die Funktion der meisten Optionen in Visual Studio geht aus ihren Beschreibungen hervor, mit Ausnahme des Kontrollkästchens  **E-Mail-Nachricht**. Verwenden Sie dieses Kontrollkästchen, wenn Sie ein Outlook-Add-In erstellen möchten, das nicht nur zusammen mit Mailelementen angezeigt wird, sondern auch mit Besprechungsanfragen, Antworten und Absagen.

Wenn Sie den Assistenten abgeschlossen haben, erstellt Visual Studio eine Projektmappe, die zwei Projekte enthält.



|**Projekt**|**Beschreibung**|
|:-----|:-----|
|Add-in-Projekt|Enthält nur eine XML-Manifestdatei, die alle Einstellungen enthält, die Ihr Add-in beschreiben. Mit diesen Einstellungen kann der Office-Host bestimmen, wann Ihr Add-in aktiviert wird und wo das Add-in angezeigt werden soll. Visual Studio generiert den Inhalt dieser Datei, sodass Sie das Projekt ausführen und Ihre Add-in sofort verwenden können. Mit dem Manifest-Editor können Sie diese Einstellungen jederzeit ändern.|
|Webanwendungsprojekt|Enthält die Inhaltsseiten Ihres Add-ins, darunter alle Dateien und Dateiverweise, die Sie zum Entwickeln von Office-fähigen HTML- und JavaScript-Seiten benötigen. Während Sie Ihr Add-in entwickeln, wird die Webanwendung von Visual Studio auf Ihrem lokalen IIS-Server gehostet. Sobald Sie Ihr Add-in veröffentlichen möchten, müssen Sie einen Server finden, der dieses Projekt hostet.Weitere Informationen zu ASP.NET-Webanwendungsprojekten finden Sie unter [ASP.NET-Webprojekte](http://msdn.microsoft.com/de-de/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## Ändern der Add-In-Einstellungen


Um die Einstellungen Ihres Add-Ins zu ändern, öffnen Sie den  **Manifest-Designer**. Der Manifest-Designer ist ein Editor, dessen Aussehen einer Eigenschaftenseite ähnelt und mit dem Sie die gebräuchlichsten Eigenschaften des Add-Ins auf eine visuellere Weise ändern können. Erweitern Sie zum Öffnen des  **Manifest-Designers** vom **Projektmappen-Explorer** den Office-Add-In-Projektknoten, wählen Sie den Ordner mit der XML-Manifestdatei aus, und drücken Sie dann die EINGABETASTE.

Zum Festlegen erweiterter Einstellungen, z. B. des Zielgebietsschemas des Add-Ins, können Sie die XML-Manifestdatei des Projekts direkt bearbeiten. Erweitern Sie im  **Projektmappen-Explorer** den Knoten des Add-In-Projekts, wählen Sie den Ordner mit der XML-Manifestdatei aus, und wählen Sie die XML-Manifestdatei. Sie können auf beliebige Elemente in der Datei zeigen, um eine QuickInfo mit einer Beschreibung des Zwecks des Elements einzublenden. Eine vollständige Liste mit Beschreibungen finden Sie unter [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md).

Der  **Manifest-Designer** umfasst mehrere Registerkarten, darunter die Registerkarte **Allgemein** mit allgemeinen Add-In-Einstellungen wie Add-In-Name und Add-In-Symbol, die Registerkarte **Aktivierung**, mit der Sie die Add-In-Anforderungen wie benötigte API-Sätze und Zielanwendungen angeben können, und die Registerkarte  **App-Domänen**, auf der Sie die Domänen von Seiten angeben können, die von Ihrem Add-In verwendet werden. Für Outlook-Add-Ins wird die Registerkarte  **Aktivierung** angezeigt, dafür jedoch zwei weitere Registerkarten: **Leseformular** und **Entwurfsvorlage**. 


### Allgemeine Einstellungen

In der folgenden Tabelle werden die Felder beschrieben, die auf der Registerkarte  **Allgemein** des Manifest-Editors angezeigt werden.



|**Eigenschaft**|**Entsprechender Wert in der XML-Manifestdatei**|**Beschreibung**|
|:-----|:-----|:-----|
|**Display Name**| `DefaultValue`-Attribut des [DisplayName](http://msdn.microsoft.com/de-de/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx)-Elements.|Name, der auf der Benutzeroberfläche der Hostanwendung angezeigt wird. Wenn ein Benutzer in Excel beispielsweise  **Einfügen**,  **Add-in** auswählt, wird dieser Name in der Liste der verfügbaren Add-ins angezeigt. Dieser Name kann auch als Titel über dem Add-in angezeigt werden.|
|**Add-in-Typ**|[OfficeApp](http://msdn.microsoft.com/de-de/library/4537b0a6-a741-332d-9e8f-4341c8b50b6a%28Office.15%29.aspx) complexType.|Typ des Add-Ins. Der Wert  **Aufgabenbereich-Add-In** weist darauf hin, dass das Add-In im Aufgabenbereich der Office-Anwendung angezeigt wird. Der Wert **Inhalts-Add-In** weist darauf hin, dass das Add-In im Textkörper eines Dokuments angezeigt wird. Der Wert **Outlook-Add-In** weist darauf hin, dass das Add-In neben einer Nachricht oder einem Termin angezeigt wird.|
|**Version**|[Version](http://msdn.microsoft.com/de-de/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx)-Element.|Gibt die Version des Add-ins an.|
|**Anbietername**|[ProviderName](http://msdn.microsoft.com/de-de/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx)-Element.|Gibt den Namen der Person oder des Unternehmens an, die bzw. das das Add-in entwickelt hat.|
|**Beschreibung**| `DefaultValue`-Attribut des [Description](http://msdn.microsoft.com/de-de/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx)-Elements.|Beschreibung, die angezeigt wird, wenn ein Benutzer in der Liste der verfügbaren Add-ins auf einen Add-in-Namen zeigt.|
|**Symbol**| `DefaultValue`-Attribut des [IconUrl](http://msdn.microsoft.com/de-de/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)-Elements.|Bild von 32 x 32 Pixel, das für Ihr Add-in im Menüband der Hostanwendung angezeigt wird.|
|**Symbol für hohe Auflösung**| `DefaultValue`-Attribut des [IconUrl](http://msdn.microsoft.com/de-de/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)-Elements.|Bild von 64 x 64 Pixel, das für Ihr Add-in im Menüband der Hostanwendung angezeigt wird.|
|**Quellspeicherort**| `DefaultValue`-Attribut des [SourceLocation](http://msdn.microsoft.com/de-de/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)-Elements.|Speicherort der ersten Seite, die im Add-in angezeigt wird, sobald es in der Hostanwendung aktiviert wird. Der Standardwert dieser Eigenschaft ist die Standard-HTML-Datei Ihres Projekts.|
|**Support-URL**| `DefaultValue`-Attribut des [SupportUrl](http://msdn.microsoft.com/de-de/library/61cff5aa-929f-7d6a-2ce9-0b92b2d6e0a7%28Office.15%29.aspx)-Elements.|Gibt die URL einer Seite an, die Supportinformationen für das Add-in enthält.|
|**Angeforderte Höhe** (nur Inhalts- und Outlook-Add-Ins)|[RequestedHeight](http://msdn.microsoft.com/de-de/library/fe949c28-a9ff-26dd-6a80-5f81abc330e8.aspx)-Element.|Anzahl der Pixel, die das Add-in, als die Höhe des Bereichs-Add-in benötigt.|
|**Angeforderte Breite** (nur Inhalts-Add-ins)|[RequestedWidth](http://msdn.microsoft.com/de-de/library/29032529-6661-fb99-1ff3-c02cc474017f.aspx)-Element.|Anzahl der Pixel, die das Add-in, als die Breite des Bereichs-Add-in benötigt.|
|**Berechtigungen**|[Permissions](http://msdn.microsoft.com/de-de/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx)-Element.|Für dieses Add-in erforderlichen Berechtigungen.|
|**Entitätshervorhebung** (nur Outlook-Add-Ins)|[DisableEntityHighlighting-Element (MailApp complexType) (App-Manifestschema v1.1)](http://msdn.microsoft.com/library/bf67a8d6-cb8f-7a58-d09d-4d5c7679d10f%28Office.15%29.aspx)-Element.|Gibt an, ob die Hervorhebung von Entitäten für dieses Outlook-Add-In deaktiviert werden soll.|

### Einstellungen der Registerkarte „Aktivierung"

Auf dieser Registerkarte können Sie die Sätze von Office JavaScript-APIs auflisten, die von einer Office-Zielanwendung unterstützt werden müssen, damit Ihr Add-in aktiviert werden kann. Wenn Ihr Add-in beispielsweise an eine Tabelle bindet, fügen Sie der Liste den TableBindings-API-Satz hinzu. Durch diese Funktion wird verhindert, dass Benutzer unschöne JavaScript-Fehler erhalten, wenn Sie unabsichtlich Ihr Add-in auf einem Office-Host aktivieren, der die in Ihrem Add-in enthaltenen Funktionen nicht unterstützt.

In der folgenden Tabelle werden die Felder beschrieben, die auf der Registerkarte  **Aktivierung** des Manifest-Editors angezeigt werden. Diese Registerkarte wird nur in Inhalts- und Aufgabenbereich-Add-ins angezeigt.



|**Eigenschaft**|**Beschreibung**|
|:-----|:-----|
|**Erforderliche API-Sätze**|Hiermit können Sie die Namen der API-Sätze und die Mindestversionen angeben, die erforderlich sind, damit Ihr Add-in ordnungsgemäß aktiviert werden kann. Um einen API-Satz hinzuzufügen, wählen Sie ihn in der Dropdownliste aus. Um einen API-Satz zu löschen, wählen Sie ihn in der Liste aus, und drücken Sie anschließend die  **ENTF**-Taste. Bei der Angabe von API-Sätzen zeigt die Seite die Office-Clients an, die die entsprechende Kombination von API-Sätzen unterstützen.
 >**Hinweis**  Da einige API-Sätze nur bestimmte Office-Clients unterstützen, verringert sich die Zahl der Office-Clients, mit denen Ihr Add-in aktiviert werden kann, je mehr API-Sätze Sie angeben.

|
|**Anwendungen**|Hiermit können Sie die Office-Anwendungen auswählen, auf die Ihr Add-in ausgerichtet sein soll. Sie können jede beliebige Office-Anwendung als Ziel festlegen, die unter Office 365 und Office 2013 SP1 verfügbar ist, oder Sie können bestimmte Office-Anwendungen als Ziel festlegen.|
|**IntelliSense**|Hiermit können Sie festlegen, ob IntelliSense Syntaxinformationen für alle Office JavaScript-APIs oder nur für APIs in der Liste  **Erforderliche API-Sätze** anzeigt.|
|**Zusammenfassung**|In diesem Abschnitt wird angezeigt, wo das Add-in basierend auf Ihren Eingaben in den Abschnitten  **Erforderliche API-Sätze** und **Anwendungen** aktiviert wird.|

### Einstellungen der Registerkarte „App-Domänen"

Die App-Domäneneinstellungen werden verwendet, wenn das Add-In mit einer entfernten Domäne kommunizieren muss (domänenübergreifende Kommunikation). In der folgenden Tabelle werden die Felder beschrieben, die auf der Registerkarte  **Add-In-Domänen** des Manifest-Editors angezeigt werden. Anhand dieser Einstellungen können Sie die Domänen angeben, die dieses Add-In zum Laden von Seiten verwenden soll.



|**Eigenschaft**|**Beschreibung**|
|:-----|:-----|
|**Geben Sie die URL einer Domäne ein**|Hier können Sie die URL einer Domäne eingeben, die von Ihrem Add-in verwendet wird. Klicken Sie auf die Schaltfläche  **Hinzufügen**, um sie in die Liste  **AppDomains** aufzunehmen.|
|**AppDomains**|Die Liste der von Ihnen angegebenen Add-in-Domänen. Sie können Elemente aus der Liste entfernen, indem Sie sie auswählen und anschließend auf die Schaltfläche  **Entfernen** klicken.|

### Einstellungen der Registerkarte „Leseformular"

Leseformulareinstellungen sind nur für Outlook-Add-Ins verfügbar. Diese Einstellungen geben an, wann ein Leseformular aktiviert wird und die zugehörigen Benutzeroberflächeneigenschaften.


|
|
|**Eigenschaft**|**Beschreibung**|
|:-----|:-----|
|Aktivierung|Gibt die Aktivierungsregeln für das Add-In für das Leseformular. Wählen Sie die entsprechende Regel oder Regeln in der Strukturansicht aus, oder fügen Sie Regeln durch Auswählen von  **Hinzufügen** aus der Dropdownliste hinzu. Optional können Sie die Nachrichtenklasse des E-Mail- oder Terminelements angeben.|
|Quellspeicherort|Gibt den Quellspeicherort für das Add-In an.|
|Angeforderte Höhe|Gibt die angeforderte Höhe in Pixeln an.|
|Aktivieren Sie das Add-In, damit es unter Elementen angezeigt wird, die auf einem Tablet-PC geöffnet sind.|Aktivieren Sie dieses Kontrollkästchen, wenn das Add-In auf einem Tablet-PC angezeigt werden soll. Wenn dieses aktiviert ist, müssen Sie auch den Quellspeicherort für den Tablet-Code und die angeforderte Höhe angeben.|
|Aktivieren Sie das Add-In, damit es unter Elementen angezeigt wird, die auf einem Smartphone geöffnet sind.|Aktivieren Sie dieses Kontrollkästchen, wenn das Add-In auf einem Smartphone angezeigt werden soll. Wenn dieses aktiviert ist, müssen Sie auch den Quellspeicherort des Smartphonecodes angeben.|

### Einstellungen der Registerkarte „Entwurfsvorlage"

Entwurfsvorlageneinstellungen sind nur für Outlook-Add-Ins verfügbar. Diese Einstellungen geben an, wann eine Entwurfsvorlage aktiviert wird und die zugehörigen Benutzeroberflächeneigenschaften.


|
|
|**Eigenschaft**|**Beschreibung**|
|:-----|:-----|
|Aktivierung|Gibt die Typen von Elementen an, die ein Entwurfsvorlagen-Add-In aktivieren. Sie können entweder eins oder beides  **E-Mail-Nachrichten** und **Termine** auswählen.|
|Quellspeicherort|Gibt den Quellspeicherort für das Add-In an.|
|Symbol für hohe Auflösung|Gibt die URL des Bilds, das für das Add-In auf einem DPI-Display mit hoher Auflösung verwendet wird. Die Auflösung muss 128x128 Pixel betragen.|
|Aktivieren Sie das Add-In, damit es unter Elementen angezeigt wird, die auf einem Tablet-PC geöffnet sind.|Aktivieren Sie dieses Kontrollkästchen, wenn das Add-In auf einem Tablet angezeigt werden soll. Wenn dieses aktiviert ist, müssen Sie auch den Quellspeicherort des Tabletcodes angeben.|
|Aktivieren Sie das Add-In, damit es unter Elementen angezeigt wird, die auf einem Smartphone geöffnet sind.|Aktivieren Sie dieses Kontrollkästchen, wenn das Add-In auf einem Smartphone angezeigt werden soll. Wenn dieses aktiviert ist, müssen Sie auch den Quellspeicherort des Smartphonecodes angeben.|

## Entwickeln des Inhalts des Add-Ins


Während das Add-in-Projekt ermöglicht, die Einstellungen zu ändern, die das Add-in beschreiben, stellt die Webanwendung den Inhalt bereit, der im Add-in angezeigt wird. 

Das Webanwendungsprojekt enthält eine Standard-HTML-Seite und eine JavaScript-Datei, die Sie als Ausgangspunkt verwenden können. Das Projekt enthält auch eine JavaScript-Datei, die allen Seiten, die Sie Ihrem Projekt hinzufügen, gemein ist. Diese Dateien sind praktisch, da sie Verweise auf andere JavaScript-Bibliotheken einschließlich der JavaScript-API für Office enthalten. 

Je ausgereifter Ihr Add-in wird, desto mehr HTML- und JavaScript-Dateien können Sie hinzufügen. Sie können die Inhalte der Standard-HTML- und JavaScript-Dateien als Beispiele für die Arten von Referenzen verwenden, die Sie anderen Seiten in Ihrem Projekt hinzufügen möchten, damit sie mit Ihrem Add-in arbeiten. In der folgenden Tabelle sind die Standard-HTML- und JavaScript-Dateien.



|**Datei**|**Beschreibung**|
|:-----|:-----|
|**Home.html**|Dies ist die Standard-HTML-Seite des Add-ins. Sie befindet sich im Ordner  **Home** des Projekts. Diese Seite wird als erste Seite im Add-in angezeigt, wenn es in einem Dokument, einer E-Mail-Nachricht oder einem Terminelement aktiviert wird. Diese Datei ist praktisch, da sie alle erforderlichen Dateiverweise enthält, um schnell loszulegen. Sobald Sie Ihr erstes Add-in erstellen, fügen Sie in diese Datei einfach Ihren HTML-Code ein.|
|**Home.js**|Dies ist die mit der Seite „Home.js" verbundene JavaScript-Datei der App. Sie befindet sich im Ordner  **Home** des Projekts. Hier können Sie jeden beliebigen Code für das Verhalten der „Home.html"-Seite in der „Home.js"-Datei eingeben. Die „Home.js"-Datei enthält Beispielcode, damit Sie loslegen können.|
|**App.js**|Dies ist die Standard-JavaScript-Datei des gesamten Add-ins. Sie befindet sich im Ordner  **Add-in** des Projekts. In diese „App.js"-Datei können Sie Code einfügen, der für das Verhalten mehrerer Seiten Ihres Add-ins normal ist. Diese „App.js"-Datei enthält Beispielcode, damit Sie loslegen können.|

 >**Hinweis**  Sie müssen diese Dateien nicht verwenden und können dem Projekt jederzeit andere Dateien hinzufügen, die Sie dann stattdessen verwenden können. Falls eine andere HTML-Datei als Startseite des Add-in angezeigt werden soll, öffnen Sie den Manifest-Editor, und verweisen Sie die  **SourceLocation**-Eigenschaft auf den Namen der Datei.


## Debuggen des Add-Ins


Wenn Sie zum Starten Ihres Add-ins bereit sind, sollten Sie den Build überprüfen und die zugehörigen Eigenschaften debuggen und anschließend die Projektmappe starten.


### Überprüfen des Builds und Debuggen der Eigenschaften

Bevor Sie die Projektmappe starten, sollten Sie sicherstellen, dass Visual Studio die gewünschte Hostanwendung öffnet. Diese Information wird auf den Eigenschaftenseiten des Projekts zusammen mit anderen Eigenschaften im Zusammenhang mit dem Erstellen und Debuggen des Add-ins angezeigt.


### So öffnen Sie die Eigenschaftenseiten eines Projekts


1. Wählen Sie imProjektmappen-Explorer den Projektnamen aus.
    
2. Klicken Sie in der Menüleiste auf  **Ansicht**,  **Eigenschaftenfenster**.
    
In der folgenden Tabelle werden die Eigenschaften des Projekts beschrieben.



|**Eigenschaft**|**Beschreibung**|
|:-----|:-----|
|**Startaktion**|Gibt an, ob Ihr Add-in in einem Office-Desktopclient oder in einem Office Online-Client im angegebenen Browser gedebuggt werden soll.|
|**Startdokument** (nur Inhalts- oder Aufgabenbereich-Add-ins)|Gibt an, welches Dokument beim Starten des Projekts geöffnet werden soll.|
|**Webprojekt**|Gibt den Namen des mit dem Add-in verbundenen Webprojekts an.|
|E-Mail-Adresse (nur Outlook-Add-Ins)|Gibt die E-Mail-Adresse des Benutzerkontos in Exchange Server oder Exchange Online an, mit dem Sie Ihr Outlook-Add-In testen möchten.|
|**EWS-URL** (nur Outlook-Add-Ins)|Die Exchange-Webdienst-URL (z. B. https://www.contoso.com/ews/exchange.aspx). |
|**OWA-URL** (nur Outlook-Add-Ins)|Outlook Web App URL (z. B. https://www.contoso.com/owa).|
|**Benutzername** (nur Outlook-Add-Ins)|Gibt den Namen Ihres Benutzerkontos in Exchange Server oder Exchange Online an.|
|**Projektdatei**|Gibt den Namen der Datei mit den Build-, Konfigurations- und anderen Informationen zum Projekt an.|
|**Projektordner**|Der Speicherort der Projektdatei.|

### Debuggen der Add-ins mithilfe eines vorhandenen Dokuments (nur Inhalts- und Aufgabenbereich-Add-ins)


Sie können dem Add-in-Projekt Dokumente hinzufügen. Wenn Sie über ein Dokument mit Testdaten verfügen, die Sie zusammen mit Ihrem Add-in verwenden möchten, öffnet Visual Studio dieses Dokument, sobald Sie das Projekt starten.


### So debuggen Sie das Add-in mithilfe eines vorhandenen Dokuments


1. Wählen Sie im  **Projektmappen-Explorer** den Add-in-Projektordner aus.
    
     >**Hinweis**  Wählen Sie das Add-in-Projekt und nicht das Webanwendungsprojekt aus.
2. Wählen Sie im Menü  **Projekt** die Option **Vorhandenes Element hinzufügen**.
    
3. Lokalisieren Sie im Dialogfeld  **Vorhandenes Element hinzufügen** das Dokument, das Sie hinzufügen möchten, und wählen Sie es aus.
    
4. Klicken Sie auf die Schaltfläche  **Hinzufügen**, um das Dokument dem Projekt hinzuzufügen.
    
5. Öffnen Sie im  **Projektmappen-Explorer** das Kontextmenü für das Projekt, und wählen Sie **Eigenschaften** aus.
    
    Die Eigenschaftenseiten des Projekts werden angezeigt.
    
6. Wählen Sie in der Liste  **Startdokument** das Dokument aus, das Sie dem Projekt hinzugefügt haben, und klicken Sie anschließend auf die Schaltfläche **OK**, um die Seite mit den Eigenschaften zu schließen.
    

### Starten der Projektmappe


Visual Studio erstellt die Projektmappe automatisch, wenn Sie sie starten. Sie können die Projektmappe über die **Menüleite** starten, indem Sie **Debuggen** und dann **Starten** auswählen.


 >**Hinweis**  Falls das Skriptdebugging in Internet Explorer nicht aktiviert ist, können Sie den Debugger nicht in Visual Studio starten. Um das Skriptdebugging zu aktivieren, öffnen Sie das Dialogfeld  **Internetoptionen**, klicken Sie auf die Registerkarte  **Erweitert**, und deaktivieren Sie dann die Kontrollkästchen  **Skriptdebugging deaktivieren (Internet Explorer)** und **Skriptdebugging deaktivieren (Andere)**.

Visual Studio erstellt das Projekt und führt die folgenden Aktionen aus.


1. Eine Kopie der XML-Manifestdatei wird erstellt und dem Verzeichnis  _Projektname_\Output hinzugefügt. Die Hostanwendung verwendet diese Kopie, wenn Sie Visual Studio starten und das Add-in debuggen.
    
2. Ein Satz von Registrierungseinträgen wird auf Ihrem Computer erstellt, die es ermöglichen, dass das Add-in in der Hostanwendung angezeigt wird.
    
3. Das Webanwendungsprojekt wird erstellt und anschließend auf dem lokalen IIS-Webserver (http://localhost) bereitgestellt. 
    
Als Nächstes führt Visual Studio die folgenden Aktionen aus:


1. Das [SourceLocation](http://msdn.microsoft.com/de-de/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)-Element der XML-Manifestdatei wird geändert, indem das Token „~remoteAppUrl" durch die vollqualifizierte Adresse der Startseite (Beispiel: http://localhost/MyApp.html) ersetzt wird.
    
2. Das Webanwendungsprojekt wird in IIS Express gestartet.
    
3. Die Hostanwendung wird geöffnet. 
    
Visual Studio zeigt beim Erstellen des Projekts keine Überprüfungsfehler im Fenster  **OUTPUT** an. Visual Studio meldet aufgetretene Fehler und Warnungen im Fenster **ERRORLIST**. Darüber hinaus weist Visual Studio im Code und im Text-Editor durch wellenförmige Unterstreichung (so genannte Wellenlinien) in unterschiedlichen Farben auf Überprüfungsfehler hin. Diese Markierungen sind ein Hinweis auf Probleme, die von Visual Studio im Code gefunden wurden. Weitere Informationen finden Sie unter [Code- und Text-Editor](http://go.microsoft.com/fwlink/?LinkID=128497). Weitere Informationen zum Aktivieren oder Deaktivieren der Validierung finden Sie in den folgenden Themen: 


- [Optionen, Text-Editor, JavaScript, IntelliSense](http://go.microsoft.com/fwlink/?LinkID=238779)
    
- [Vorgehensweise: Festlegen von Validierungsoptionen für die HTML-Bearbeitung in Visual Web Developer](http://msdn.microsoft.com/de-de/library/vstudio/0byxkfet%28v=vs.100%29.aspx)
    
- [Validierung, CSS, Text-Editor, Dialogfeld "Optionen"](http://go.microsoft.com/fwlink/?LinkID=238780)
    
Informationen zu den Überprüfungsregeln der XML-Manifestdatei in Ihrem Projekt finden Sie unter [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md).


### Anzeigen eines Add-ins in Excel, Word oder Project, und schrittweises Durchgehen des Codes


Wenn Sie die Eigenschaft  **Startdokument** des Add-in-Projekts auf Excel oder Word festlegen, erstellt Visual Studio ein neues Dokument, und die App wird angezeigt. Wenn Sie die Eigenschaft **Startaktion** des App-Projekts auf ein vorhandenes Dokument festlegen, öffnet Visual Studio zwar das Dokument, doch Sie müssen dann die App manuell einfügen. Wenn Sie die Eigenschaft **Startdokument** auf **Microsoft Project** festlegen, müssen Sie die App ebenfalls manuell hinzufügen.


### So zeigen Sie eine Office-Add-In in Excel oder Word an


1. Wählen Sie in Excel oder Word auf der Registerkarte  **Einfügen** die Option **Office-Add-ins** aus.
    
2. Wählen Sie in der eingeblendeten Liste Ihr Add-in aus.
    

### So zeigen Sie eine Office-Add-In in Project an


1. Wählen Sie in Project auf der Registerkarte  **Projekt** die Option **Office-Add-Ins** aus.
    
2. Wählen Sie in der eingeblendeten Liste Ihr Add-in aus.
    
In Visual Studio können Sie dann Haltepunkte erstellen. Hiermit können Sie anschließend mit Ihrem Add-in interagieren und den Code in Ihren HTML-, JavaScript- und C#- oder VB-Codedateien durchlaufen.


### Anzeigen des Outlook-Add-Ins in Outlook und Durchlaufen des Codes


Öffnen Sie eine E-Mail-Nachricht oder ein Terminelement, um das Add-in in Outlook anzuzeigen.

Outlook aktiviert die Add-In für das Element, vorausgesetzt die Aktivierungskriterien sind erfüllt. Die Add-In-Leiste wird oben im Inspektorfenster oder im Lesebereich angezeigt, und Ihr Outlook-Add-In wird als Schaltfläche in der Add-In-Leiste angezeigt. Wenn Ihr Add-In über ein Add-In-Befehl verfügt, wird eine Schaltfläche im Menüband angezeigt, entweder die Standardregisterkarte oder eine benutzerdefinierte Registerkarte, und das Add-In wird nicht in der Add-In-Leiste angezeigt.

Wählen Sie die Schaltfläche für Ihr Outlook-Add-In aus, um sie anzuzeigen.

In Visual Studio können Sie dann Haltepunkte erstellen. Hiermit können Sie anschließend mit Ihrem Outlook-Add-In interagieren und den Code in Ihren HTML-, JavaScript- und C#- oder VB-Codedateien durchlaufen. 

Darüber hinaus können Sie Ihren Code ändern und die Auswirkungen dieser Änderungen in Ihrem Outlook-Add-In überprüfen, ohne dass Sie die Office-Add-In schließen und das Projekt erneut starten müssen. Öffnen Sie in Outlook einfach das Kontextmenü für das Outlook-Add-In, und wählen Sie dann  **Erneut laden** aus.


### Ändern von Code und Fortfahren mit dem Debuggen des Add-Ins, ohne das Projekt erneut starten zu müssen


Sie können Ihren Code ändern und die Auswirkungen dieser Änderungen im Add-in anzeigen, ohne dass Sie die Hostanwendung schließen und das Projekt neu starten müssen. Öffnen Sie nach dem Ändern des Codes das Kontextmenü des Add-ins, und klicken Sie auf  **Neu laden**. Wenn Sie das Add-in neu laden, wird die Verbindung zum Visual Studio-Debugger getrennt. Daher können Sie die Auswirkungen der Änderung anzeigen, aber Sie können den Code nicht erneut Schritt für Schritt durchlaufen. Dazu müssen Sie den Visual Studio-Debugger allen verfügbaren Iexplore.exe-Prozessen zuordnen.


### So ordnen Sie den Visual Studio-Debugger allen verfügbaren Iexplore.exe-Prozessen zu


1. Wählen Sie in Visual Studio die Option  **DEBUGGEN** und dann **An den Prozess anhängen** aus.
    
2. Wählen Sie im Dialogfeld  **An den Prozess anhängen** alle verfügbaren **Iexplore.exe**-Prozesse aus, und wählen Sie dann die Schaltfläche  **An den Prozess anhängen**.
    

## Nächste Schritte



- [Erstellen eines Aufgabenbereich- oder Inhalts-Add-Ins mit Visual Studio](create-a-task-pane-or-content-add-in-with-visual-studio.md)
    
- [Verpacken des Add-Ins mit Napa oder Visual Studio für die Veröffentlichung](../publish/package-your-add-in-using-napa-or-visual-studio.md)
    
