
# Interaktionsmuster für Office-Add-Ins
Informationen zu den Szenarien für gängige UX-Interaktionsmuster in Inhalts-, Aufgabenbereich- und Outlook-Add-Ins

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Office-Add-Ins können Dokumenterstellungs- und Produktivitätsumgebungen optimieren sowie Inhalte in Office-Hostanwendungen mit größeren webbasierten Workflows verbinden. Es gibt eine Reihe gängiger Szenarien für Inhalts-, Aufgabenbereich- und Outlook-Add-Ins, die Sie möglicherweise entwickeln. In diesem Artikel werden einige der gängigsten Szenarien beschrieben und empfohlene Interaktionsmuster für die Add-In-UX. Sie können diese Interaktionsmuster je nach Ihren einzigartigen Szenarien unterteilen, kombinieren oder beliebig mischen.

 **Gängige Add-In-Szenarien**



|**Add-In-Typ**|**Gängige Szenarien**|
|:-----|:-----|
|Inhalt|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Visualisieren von Daten</p></li><li><p>Widgets und Tools</p></li></ul>|
|Aufgabenbereich|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Umwandeln und Verarbeiten von Daten</p></li><li><p>Effektivere und effizientere Dokumenterstellung</p></li><li><p>Suchen nach Inhalten und Einfügen von Medien</p></li><li><p>Veröffentlichen oder Hochladen von Inhalten in einen Webdienst</p></li></ul>|
|Outlook|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Bridging zwischen E-Mail-Inhalten und einer externen Anwendung</p></li><li><p>Vermitteln weiterer Informationen über die Inhalte in einer E-Mail-Nachricht oder einen Termin</p></li><li><p>Bereitstellen von Informationen, die Ihnen zu mehr Produktivität verhelfen</p></li></ul>|

## Visualisieren von Daten mit einem Inhalts-Add-In


Dieses Beispiel zeigt ein Inhalts-Add-in für Excel zum Generieren eines Diagramms anhand von Daten in einem Tabellenblatt.

In diesem Interaktionsmuster wird das Add-in erst aktiv, nachdem Sie Daten zum Generieren des Diagramms ausgewählt und an das Add-in gebunden haben. Es ist wichtig, dass der Zweck des Add-ins und die Schritte für die Aktivierung in der Ausgangsansicht des Add-ins ersichtlich sind. 


|||
|:-----|:-----|
|
**Inhalts-Add-in für Excel zum Generieren eines Diagramms anhand von Daten in einem Tabellenblatt**
![Inhalts-App für Excel zum Generieren eines Diagramms anhand von Daten in einem Tabellenblatt](../../../images/off15appUXFig01.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Um zu unterstreichen, dass vor Auswahl der Schaltfläche eine Aktion ausgeführt werden muss, zeigen Sie Anweisungen zusammen einer Schaltfläche vom Typ "Deaktiviert" (A) an.</p></li><li><p>Nachdem Sie einen Zellbereich ausgewählt haben, wird die Schaltfläche <span class="ui">Diagramm erstellen</span> aktiv (B – C).</p></li><li><p>Die Visualisierung füllt den Container aus und ersetzt die vorherige Ansicht (D).</p></li><li><p>Zeigen Sie ggf. zusätzliche Benutzeroberflächenelemente am unteren Rand des Add-ins zusammen mit einer Einstellungsschaltfläche (Zahnrad) an, um zu einer Ansicht zu führen, in der Sie das Add-in zurücksetzen oder verwalten können.</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins, bei denen Sie vor der Aktivierung Daten auswählen müssen</p></li></ul>|

## Umwandeln von Inhalten mit einem Aufgabenbereich-Add-In


Dieses Beispiel zeigt ein Aufgabenbereich-Add-in, das Text in Ihrem Dokument in eine andere Sprache übersetzt.

In diesem Interaktionsmuster müssen Sie zuerst den zu übersetzenden Text im Dokument auswählen.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in, das Text in Ihrem Dokument in eine andere Sprache übersetzt**
![Aufgabenbereich-App, die Text in Ihrem Dokument in eine andere Sprache übersetzt](../../../images/off15appUXFig02.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Vermitteln Sie den Zweck des Add-ins in einer Überschrift, und weisen Sie darauf hin, dass Sie zuerst eine Auswahl treffen müssen (A).</p></li><li><p>Das Sprachmenü und die Schaltfläche <span class="ui">Translate</span> (Übersetzen) sind deaktiviert. Dies unterstreicht, dass Sie zunächst eine Aktion ausführen müssen, bevor Sie fortfahren können. Nachdem Sie Inhalte im Dokument ausgewählt haben, werden diese beiden Elemente aktiv (D).</p></li><li><p>Nachdem Sie <span class="ui">Translate</span> (Übersetzen) ausgewählt haben, wird die Benutzeroberfläche eingeblendet. Sie enthält die übersetzten Inhalte sowie eine Schaltfläche, mit der diese wieder in das Dokument eingefügt werden können (E).</p></li><li><p>Sie können eine Schaltfläche vom Typ <span class="ui">Löschen</span> oder <span class="ui">Zurücksetzen</span> bereitstellen, mit der Sie wieder zur Erstansicht gelangen.</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins, bei denen Sie vor der Aktivierung Daten auswählen müssen</p></li><li><p>Benutzeroberflächenelemente, die eingeblendet oder angezeigt werden, während Sie ein Szenario durchlaufen</p></li></ul>|

## Verarbeiten von Daten mit einem Aufgabenbereich-Add-In


Dieses Beispiel zeigt ein Aufgabenbereich-Add-in, das Daten in Excel überprüft.

In diesem Interaktionsmuster müssen Sie einen zuerst einen Zellbereich im Tabellenblatt auswählen.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in, das Daten in Excel überprüft**
![Aufgabenbereich-App, die Daten in Excel überprüft](../../../images/off15appUXFig03.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Der Zweck des Add-ins wird in der Überschrift beschrieben. Anweisungen helfen Ihnen beim Einstieg.</p></li><li><p>Die Schaltfläche <span class="ui">Send selected data</span> (Ausgewählte Daten senden) ist deaktiviert. Dies unterstreicht, dass Sie eine Aktion ausführen müssen, um fortzufahren (A).</p></li><li><p>Nachdem Sie einen Zellbereich im Tabellenblatt (B) ausgewählt haben, wird die Schaltfläche <span class="ui">Send selected data</span> (Ausgewählte Daten senden) aktiviert.</p></li><li><p>Nachdem Sie diese Schaltfläche ausgewählt haben, werden die Benutzeroberflächenelemente durch den ausgewählten Zellbereich, eine Statusleiste und die Schaltfläche <span class="ui">Cancel</span> (Abbrechen) ersetzt.</p></li><li><p>Die Statusleiste zeigt den Status des Prozesses an, und die Schaltfläche <span class="ui">Cancel</span> (Abbrechen) ermöglicht dessen Unterbrechung (D).</p></li><li><p>Nach Abschluss des Prozesses werden die Ergebnisse automatisch angezeigt (E). Durch das Auswählen eines Elements in der Liste wird die entsprechende Zelle in der Tabelle aktiviert.</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Prozesse, die unbestimmte Zeit dauern</p></li></ul>|

## Analysieren von Inhalten mit einem Aufgabenbereich-Add-In


Dieses Beispiel zeigt ein Aufgabenbereich-Add-in, das während der Benutzereingabe Wortdefinitionen anzeigt.

In diesem Interaktionsmuster müssen Sie zuerst Text im Dokument auswählen, um Ergebnisse zu erhalten.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in, das während der Benutzereingabe Wortdefinitionen anzeigt**
![Aufgabenbereich-App, die während der Benutzereingabe Wortdefinitionen anzeigt](../../../images/off15appUXFig04.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>In einer Überschrift werden der Zweck des Add-ins und die ersten Schritte erklärt (A).</p></li><li><p>Die automatische Suche ist standardmäßig aktiviert, und es gibt eine Option zum Deaktivieren (B).</p></li><li><p>Nachdem Sie eine Auswahl getroffen haben, zeigt das Add-in die entsprechenden Inhalte an (D).</p></li><li><p>Stellen Sie einen Link zu weiteren Informationen zur Verfügung (E).</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins, die bei der Dokumenterstellung automatisch Inhalte zurückgeben</p></li><li><p>Add-ins, bei denen Sie vor der Aktivierung Inhalte auswählen müssen</p></li></ul>|

## Suchen nach Inhalten mit einem Aufgabenbereich-Add-In


Dieses Beispiel zeigt ein Aufgabenbereich-Add-in zum Suchen von Inhalten.

In diesem Interaktionsmuster geben Sie zuerst eine Zeichenfolge in das Suchfeld ein oder treffen Ihre Auswahl aus einer Liste wichtiger Inhalte.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in zum Suchen von Inhalten**
![Aufgabenbereich-App zum Suchen von Inhalten](../../../images/off15appUXFig05.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Die Ausgangsansicht enthält ein Suchfeld<span class="ui"> </span>(A) und eine Liste wichtiger Inhalte (B).</p></li><li><p>Wenn Sie eine Zeichenfolge in das Suchfeld eingeben, wird das Suchsymbol durch ein Schließsymbol (C) ersetzt.</p></li><li><p>Durch Klicken auf das Schließsymbol gelangen Sie wieder zur Ausgangsansicht.</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins, die bei der Dokumenterstellung automatisch Inhalte zurückgeben</p></li><li><p>Add-ins, bei denen Sie vor der Aktivierung Inhalte auswählen müssen</p></li></ul>|

## Einfügen von Medien mit einem Aufgabenbereich-Add-In


In diesem Interaktionsmuster können Sie ein Bild in den Suchergebnissen auswählen, um es in Ihr Dokument einzufügen.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in für das Einfügen eines Bilds**
![Aufgabenbereich-App zum Einfügen eines Bildes](../../../images/off15appUXFig06.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Sie haben eine Liste von Suchergebnissen gefiltert (A) und Inhalte zum Einfügen ausgewählt (B).</p></li><li><p>Eine Detailansicht des ausgewählten Inhalts (C) wird mit einer Schaltfläche angezeigt, über die Sie zurück zur Liste wechseln können.</p></li><li><p>Die Schaltfläche <span class="ui">Insert Photo</span> (Foto einfügen) befindet sich in der Fußleiste (D). Nachdem Sie diese Schaltfläche ausgewählt haben, wird das Bild in das Dokument eingefügt.</p></li><li><p>Zum eingefügten Inhalt gehört eine kurze Beschreibung der Quelle des Bilds (E). </p></li><li><p>Benutzeroberflächenelemente im Add-in bestätigen visuell den Erfolg der Aktion.</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins zum Einfügen von Inhalten</p></li></ul>|

## Einfügen von ausgewähltem Text mit einem Aufgabenbereich-Add-In


In diesem Interaktionsmuster wählen Sie Text aus den Suchergebnissen aus, um diesen in Ihr Dokument einzufügen.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in zum Einfügen von Text**
![Aufgabenbereich-App zum Einfügen von Text](../../../images/off15appUXFig07.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Sie haben bereits ein Inhaltselement gefunden (A).</p></li><li><p>Die deaktivierte Schaltfläche <span class="ui">Insert Selection</span> (Auswahl einfügen) wird in der Fußzeile angezeigt (B).</p></li><li><p>Wenn Sie eine Zeichenfolge auswählen (C), wird die Schaltfläche <span class="ui">Insert Selection</span> (Auswahl einfügen) aktiv.</p></li><li><p>Nachdem Sie auf diese Schaltfläche geklickt haben, fügt das Add-in den ausgewählten Text mit einem Hinweis auf die Inhaltsquelle in das Dokument ein (E).</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins zum Durchführen von Recherchen und Einfügen von Inhalten</p></li></ul>|

## Veröffentlichen von Inhalten für einen Webdienst mit einem Aufgabenbereich-Add-In


Dieses Beispiel zeigt ein Aufgabenbereich-Add-in zum Veröffentlichen eines Dokuments als Blogbeitrag

In diesem Interaktionsmuster haben Sie Inhalte in einem Dokument erstellt, die Sie in Ihrem Blog veröffentlichen möchten.


|||
|:-----|:-----|
|
**Aufgabenbereich-Add-in zum Veröffentlichen eines Dokuments als Blogbeitrag**
![Aufgabenbereich-App zum Veröffentlichen eines Dokuments als Blogbeitrag](../../../images/off15appUXFig08.png)|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Zuerst wird ein Anmeldeformular angezeigt, in das Sie Ihre Anmeldeinformationen eingeben müssen (A).</p></li><li><p>Es werden Links zu Informationen zum Erstellen eines Kontos und Beheben häufiger Anmeldefehler zur Verfügung gestellt (B). Durch Klicken auf diese Links werden die Seiten in einem Browser geöffnet.</p></li><li><p>Nachdem Sie sich angemeldet haben, zeigt das Add-in ein Formular zum Erstellen eines neuen Blogbeitrags an (C).</p></li><li><p>Der Name des Kontos, bei dem Sie angemeldet sind (und unter dem Sie den Beitrag erstellen) wird oben im Add-in angezeigt. Das Add-in stellt einen Link zu einer Vorschau des Beitrags zur Verfügung (D). Wenn Sie auf diesen Link klicken, wird eine Vorschau im Browser angezeigt.</p></li><li><p>Nachdem Sie auf <span class="ui">Create post</span> (Beitrag erstellen) geklickt haben, zeigt das Add-in eine Ansicht an, die bestätigt, dass der Dokumentinhalt veröffentlicht wurde (E).</p></li><li><p>Das Add-in stellt einen Link zum Anzeigen des Beitrags in einem Browser (F) sowie eine Schaltfläche zum Erstellen eines weiteren Beitrags an (G).</p></li></ul>Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins für die Ausgabe, Veröffentlichung oder Freigabe von Inhalten in sozialen Netzwerken, auf Blogwebsites oder in Webdiensten</p></li><li><p>Add-ins, die erfordern, dass Sie sich bei einem Dienst anmelden</p></li></ul>|

## Erhalten weiterer Informationen über Personen mit einem Mail-Add-In


 **Beispiel 1**


|||
|:-----|:-----|
|
**Ergebnis- und Detailseite**
![Ergebnis- und Detailseite](../../../images/off15appUXFig09.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Das Offenlegen des Umfangs Ihrer Inhalte, wenn Sie über große Datenmengen verfügen, die Sie präsentieren möchten</p></li><li><p>Detailseiten, die die vollständige Größe des Add-in-Containers nutzen</p></li><li><p>Navigationsmodelle, die von einem Informationsfluss in beide Richtungen profitieren</p></li></ul>|
 **Beispiel 2**


|||
|:-----|:-----|
|
**Detailseite mit beständiger Navigation**
![Detailseite mit beständiger Navigation](../../../images/off15appUXFig10.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Das standardmäßige Anzeigen des ersten Ergebnisses eines Datasets</p></li><li><p>Navigationsstrukturen, die wie Registerkarten funktionieren (lineare Navigation auf einer Ebene)</p></li><li><p>Verringern von Auswahlaktionen dadurch, dass Navigation jederzeit möglich ist</p></li><li><p>Das Bereitstellen von Platz, um Navigationselemente jederzeit anzeigen zu können</p></li></ul>|

## Erhalten weiterer Informationen über Personen mit einem Mail-Add-In


 **Beispiel 1**


|||
|:-----|:-----|
|
**Ergebnis- und Detailseite**
![Ergebnis- und Detailseite](../../../images/off15appUXFig11.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Das Offenlegen des Umfangs Ihrer Inhalte, wenn Sie über große Datenmengen verfügen, die Sie zeigen möchten</p></li><li><p>Die Aufforderung zum Treffen einer Auswahl, bevor weitere Details angezeigt werden</p></li><li><p>Detailseiten, die die vollständige Größe des Add-in-Containers nutzen</p></li><li><p>Navigationsmodelle, die von einem Informationsfluss in beide Richtungen profitieren</p></li></ul>|
 **Beispiel 2**


|||
|:-----|:-----|
|
**Detailseite mit sekundären Inhalten**
![Detailseite mit sekundärem Inhalt](../../../images/off15appUXFig12.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Fälle, bei denen Sie ein Inhaltselement hervorheben möchten</p></li><li><p>Das Verfügbarmachen Ihrer Inhalte ohne Interaktion mit dem Benutzer</p></li><li><p>Beständige Navigation (kann diesem Modell für eine Kombination aus Einfachheit und Navigationserleichterung hinzugefügt werden)</p></li></ul>|

## Verbinden mit einem Onlinedienst und Präsentieren von Daten


Diese Beispiele zeigen Interaktionsmuster zum Abrufen von Daten und Inhalten von einem Onlinedienst. Sie können in allen drei Add-in-Typen verwendet werden: Inhalts-Add-Ins, Aufgabenbereich-Add-Ins und Outlook-Add-Ins.

 **Beispiel 1**


|||
|:-----|:-----|
|
**Karussell**
![Karussell](../../../images/off15appUXFig13.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Kleine Datenmengen, die nacheinander oder in Gruppen zur Verfügung gestellt werden können</p></li><li><p>Das Verfügbarmachen von Inhalten in einem Katalogformat, z. B. Bildschirmpräsentationen oder Bildkataloge</p></li><li><p>Wenn ein Navigationsmodell vom Typ "Weiter/Zurück" gut funktioniert</p></li></ul>|
 **Beispiel 2**


|||
|:-----|:-----|
|
**Assistent**
![Assistent](../../../images/off15appUXFig14.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Inhalte, die in einer bestimmten Reihenfolge präsentiert werden müssen</p></li><li><p>Große Inhaltsmengen, die am besten in einer Folge kleiner Abschnitte präsentiert werden</p></li><li><p>Mit dem Lesen eines Buchs vergleichbare Umgebungen</p></li><li><p>Wenn für das Erledigen einer Aufgabe eine Reihe von Schritten oder Aktionen erforderlich ist</p></li></ul>|
 **Beispiel  3**


|||
|:-----|:-----|
|
**Formular, Ergebnisse und Details**
![Formular, Ergebnisse und Details](../../../images/off15appUXFig15.jpg)|Am besten geeignet für:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Add-ins, die Dateneingaben erfordern</p></li><li><p>Einstiegspunkt in ein Ergebnis- und Detailmuster</p></li></ul>|

## Zusätzliche Ressourcen



- [Designrichtlinien für Office-Add-Ins](../add-in-design.md)
    
