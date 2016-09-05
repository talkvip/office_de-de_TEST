
# Bewährte Methoden für die Entwicklung von Office-Add-Ins


Effektive Add-Ins bieten eine einzigartige und überzeugende Funktionalität, die Office-Anwendungen auf visuell ansprechende Weise erweitert. Stellen Sie eine ansprechende Benutzeroberfläche für die Benutzer zum Erstellen eines großartigen Add-Ins bereit, entwerfen Sie eine erstklassigen Benutzeroberfläche und optimieren Sie die Add-In-Leistung. Wenden Sie die in diesem Artikel beschriebenen bewährten Methoden zum Erstellen von Add-Ins an, mit denen die Benutzer ihre Aufgaben schnell und effizient durchführen können.

## Angeben eines klaren Nutzens



- Erstellen Sie Add-Ins, mit denen Benutzer Aufgaben schnell und effizient erledigen können. Konzentrieren Sie sich auf Szenarios, die für Office-Anwendungen sinnvoll sind. Beispiel:
 - Erledigen Sie wichtige Aufhaben schneller und einfacher und mit weniger Unterbrechungen.
 - Ermöglichen Sie neue Szenarios in Office.
 - Betten Sie ergänzende Dienste in Office-Hosts ein.
 - Verbessern Sie die Office-Benutzeroberfläche, um die Produktivität zu erhöhen.
- Stellen Sie sicher, dass der Nutzen des Add-Ins für Benutzer sofort ersichtlich ist, indem Sie ihnen ein [ansprechendes Erlebnis bei der ersten Ausführung ](#create-an-engaging-first-run-experience) bieten.
- Erstellen Sie einen [effektiven Office Store-Eintrag ](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx). Stellen Sie die Vorteile des Add-Ins im Titel und in der Beschreibung klar dar. Verlassen Sie sich nicht auf Ihre Marke, um die Funktionsweise Ihres Add-Ins zu kommunizieren.


## Erstellen Sie einer ansprechenden Benutzeroberfläche



- Locken Sie neue Benutzer mit einer einfach zu verwendenden und intuitiven Benutzeroberfläche. Beachten Sie, dass Benutzer noch beim Herunterladen des Add-Ins entscheiden, ob Sie es überhaupt verwenden sollen.

 - Kommunizieren Sie deutlich die Schritte, die die Benutzer ausführen müssen, um Ihr Add-In zu verwenden. Verwenden Sie Videos, Unterlagen, Seitenbereiche und andere Ressourcen, um Benutzer anzulocken.

 - Verstärken Sie das Werteversprechen des Add-Ins beim Start anstatt die Benutzer lediglich zum Anmelden aufzufordern.

 - Stellen Sie für Benutzer eine anpassbare Benutzeroberfläche bereit.

    ![Ein Screenshot, der einen Add-In-Aufgabenbereich mit ersten Schritten neben einem Add-In ohne erste Schritte zeigt](../../images/586202ad-333b-417c-ad31-cc6eb952b239.png)

  - Wenn das Inhalts-Add-in an Daten im Dokument des Benutzers gebunden ist, fügen Sie Beispieldaten oder eine Vorlage ein, damit Benutzer das zu verwendende Datenformat sehen können.

    ![Ein Screenshot, der ein Inhalts-Add-In mit Daten neben einem Inhalts-Add-In ohne Daten zeigt.](../../images/7de2215f-ccef-4f82-aa9d-babcbddae0c6.png)

- Bieten Sie [kostenlose Testversionen an](http://msdn.microsoft.com/library/145d9466-3c3d-4294-aa23-82068a8e7ae9.aspx%28Office.15%29.aspx#sectionSection1). Wenn für Ihr Add-In ein Abonnement erforderlich ist, stellen Sie einige Funktionen ohne Abonnement zur Verfügung.

- Einfacher Registrierungsvorgang. Sorgen Sie dafür, dass die Informationen im Vorfeld eingefügt werden (E-Mail, Anzeigename), und überspringen Sie die E-Mail-Verifizierung.

- Vermeiden von Popups. Wenn Sie diese verwenden müssen, führen Sie den Benutzer durch das Aktivieren Ihres Popups.

- Verwenden Sie [die Authentifizierung mithilfe von einmaligem Anmelden (SSO-Authentifizierung)](../outlook/authenticate-a-user-with-an-identity-token.md).

Für Vorlagen mit Mustern, die Sie beim Entwickeln Ihrer Anpassung für das erste Ausführen anwenden können, siehe  [UX-Entwurfsmuster für Office-Add-Ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## Verwenden von Add-In-Befehlen

- Stellen Sie relevante UI-Eingabepunkte für das Add-In unter Verwendung von [Add-In-Befehlen ](../design/add-in-commands.md) bereit.

- Verwenden Sie Befehle, die eine bestimmte Aktion mit einem klaren Ergebnis für den Benutzer darstellen. Mehrere Aktionen in einer einzelnen Schaltfläche sollten nicht kombiniert werden.

- Stellen Sie präzise Aktionen bereit, mit denen allgemeine Aufgaben im Add-In effizienter durchzuführen sind. Minimieren Sie die Anzahl der Schritte, die notwendig sind, um eine Aktion auszuführen.

- Für Add-Ins, die das Office-Menüband erweitern:
    - Platzieren Sie Befehle in bestehende Registerkarten (Einfügen, Überprüfung usw.), wenn die angebotene Funktion dazu passt. Wenn das Add-In Benutzern bspw. das Einfügen von Medien ermöglicht, fügen Sie eine Gruppe zur Registerkarte „Einfügen“ hinzu. Beachten Sie, dass nicht alle Registerkarten in allen Office-Versionen zur Verfügung stehen. Weitere Informationen finden Sie unter [XML-Manifest für Office-Add-Ins](../overview/add-in-manifests.md). 
    - Platzieren Sie Befehle in die Registerkarte „Start“, wenn die Funktion nicht zu einer anderen Registerkarte passt und Sie weniger als sechs Befehle der obersten Ebene haben. Sie können Befehle auch zur Registerkarte „Start“ hinzufügen, wenn Ihr Add-In in verschiedenen Office-Versionen (z. B. Office Desktop und Office Online) funktionieren muss und eine bestimmte Registerkarte nicht in allen Versionen zur Verfügung steht (z. B. ist die Registerkarte „Entwurf“ nicht in Office Online vorhanden).  
    - Platzieren Sie die Befehle in einer benutzerdefinierten Registerkarte, wenn Sie mehr als sechs Befehle der obersten Ebene haben. 
  - Benennen Sie Ihre Gruppe so, dass sie dem Namen Ihres Add-Ins entspricht. Wenn Sie über mehrere Gruppen verfügen, benennen Sie jede Gruppe basierend auf der Funktion, die die Befehle in dieser Gruppe bieten.
  - Verwenden Sie keine überflüssigen Schaltflächen zu Ihrem Add-In hinzu.

     >
  **Note**  Add-ins that take up too much space might not pass [Office Store validation](https://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe(Office.15).aspx).

- Für alle Symbole:
    - Verwenden Sie aussagekräftige Symbole und [Bezeichnungen](http://msdn.microsoft.com/library/8cef4fce-e6a1-459b-951f-47ac03ec95a6%28Office.15%29.aspx) für Schaltflächen, die die Aktion des Benutzers präzise beschreiben.


 - Verwenden Sie das PNG-Format mit einem transparenten Hintergrund.

 - Schließen Sie [alle acht unterstützten Größen](https://msdn.microsoft.com/EN-US/library/mt267547.aspx#bk_resources) mit ein. Dadurch kann die beste Lösung für alle unterstützten Auflösungen erstellt werden.

  - Verwenden Sie das visuelle Office-Layout. Beispiel:

    - Halten Sie die Formen einfach, und vermeiden Sie mehrere Farben. Komplexe Grafiken sind bei kleineren Größen und Auflösungen schwierig zu erkennen.

    - Verwenden Sie visuelle Metaphern nicht mehrfach für unterschiedliche Befehle. Das gleiche Symbol für verschiedene Aktionen führt zu Verwirrung.

    - Beschriften Sie Ihre Schaltflächen klar und prägnant. Verwenden Sie eine Kombination aus visuellen Informationen und Textinformationen, um die Bedeutung zu vermiteln.

    - Testen Sie Ihre Symbole in hellen und dunklen Office-Designs und bei kontrastreichen Einstellungen. Beachten Sie, dass Symbole auf dunklen Hintergründen oder bei hohem Kontrat vielleicht schwer zu erkennen sind.

    - Verwenden Sie konsistente Symbolgrößen und -positionen, um die visuelle Ausrichtung am Menüband zu unterstützen.


    ![Ein Screenshot, der Add-In-Befehlsschaltflächen, die dem Office-Stil entsprechen, neben Schaltflächen zeigt, die dem Stil nicht entsprechen](../../images/31e11214-61e8-41c1-888c-29d167cb9486.png)


- Bieten Sie auch eine Version Ihres Add-Ins an, die auch auf Hosts funktioniert, die Befehle nicht unterstützen. Ein einzelnes Add-In-Manifest kann sowohl in Hosts funktionieren, die Befehle unterstützen, als auch in solchen, die dies nicht tun (mit Aufgabenbereichen).

    ![Ein Screenshot, der ein Aufgabenbereich-Add-In in Office 2013 sowie dasselbe Add-In zeigt, das Add-In-Befehle in Office 2016 verwendet](../../images/4f90a3cc-8cc4-4879-9a03-0bb2b6079026.png)



## Anwenden von UX-Entwurfsprinzipien



- Stellen Sie sicher, dass das Design und die Funktionalität des Add-Ins die Office-Benutzeroberfläche ergänzt. Verwenden Sie [Office UI Fabric](https://dev.office.com/fabric).

- Ziehen Sie Inhalt Chrom vor. Vermeiden Sie überflüssige Benutzeroberflächenelemente, die keinen Nutzen für die Benutzeroberfläche mit sich bringen.

- Überlassen Sie Benutzern das Steuern. Stellen Sie sicher, dass Benutzer wichtige Entscheidungen verstehen und Aktionen auf einfache Weise rückgängig machen können, die das Add-In ausführt.

- Setzen Sie Branding ein, um das Vertrauen der Benutzer zu gewinnen und sie zu lenken. Verwenden Sie Branding nicht dazu, Benutzer zu verwirren oder zu Werbezwecken.

- Vermeiden Sie das Scrollen. Optimieren Sie das Design für die Auflösung von 1366x768.

- Verwenden Sie keine nicht lizenzierten Bilder.

- Verwenden Sie eine [klare und einfache Sprache](http://msdn.microsoft.com/library/8cef4fce-e6a1-459b-951f-47ac03ec95a6%28Office.15%29.aspx) im Add-In.

- 
  [Eingabehilfen](http://msdn.microsoft.com/library/3be1abbb-237a-48ec-8e17-72caa25a3cb2%28Office.15%29.aspx). Ermöglichen Sie allen Benutzern eine einfache Interaktion mit Ihrem Add-In, und integrieren Sie Eingabehilfstechnologien wie die Sprachausgabe.

- Entwerfen Sie für alle Plattformen und Eingabemethoden, einschließlich Tastatur, Maus und [Toucheingabe](#toucheingabe). Stellen Sie sicher, dass Ihre Benutzeroberfläche auf verschiedenen Formfaktoren reagiert.

Informationen zu Vorlagen mit Entwurfsprinzipien, die Sie beim Entwickeln des Add-Ins verwenden und anpassen können, finden Sie unter [UX-Entwurfsmuster für Office-Add-Ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

### Optimierung für Toucheingabe



- Verwenden Sie die [Context.touchEnabled ](../../reference/shared/office.context.touchenabled.md)-Eigenschaft, um zu ermitteln, ob in der Hostanwendung, die das Add-In ausführt, Toucheingabe zulässig ist.

     >**Hinweis**  Diese Eigenschaft wird in Outlook nicht unterstützt.
- Stellen Sie sicher, dass alle Steuerelemente für die Toucheingabe entsprechend angepasst sind. Schaltflächen haben beispielsweise ausreichende Touchziele, und Eingabefelder sind groß genug.

- Vertrauen Sie nicht auf Nicht-Touchmethoden wie darauf zeigen oder mit der rechten Maustaste klicken.

- Stellen Sie sicher, dass das Add-In im Hoch- und Querformat funktioniert. Beachten Sie, dass auf Touchgeräten ein Teil des Add-Ins aufgrund der Bildschirmtastatur möglicherweise ausgeblendet ist.

- Testen Sie das Add-in auf einem echten Gerät mit [Querladen](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).


 >**Hinweis**  Bei Verwendung von [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) für Entwurfselemente, werden dort viele dieser Aspekte behandelt.


## Optimieren und Überwachen der Add-In-Leistung



- Schnelle Benutzeroberflächenantworten. Das Add-In sollte in 500 ms oder weniger geladen werden.

- Stellen Sie sicher, dass innerhalb von weniger als einer Sekunde eine Reaktion auf alle Benutzerinteraktionen folgt.

-  Stellen Sie Ladeanzeigen bei lang andauernden Vorgängen bereit.

- Verwenden Sie ein CDN für Hostimages, Ressourcen und Bibliotheken. Laden Sie so viel wie möglich von einem zentralen Ort.

- Befolgen Sie die standardmäßigen Webmethoden zur Optimierung Ihrer Webseite. Verwenden Sie in der Produktion nur kleinere Versionen der Bibliotheken. Laden Sie nur Ressourcen, die Sie benötigen, und optimieren Sie die Ladevorgänge dieser Ressourcen.

- Wenn Vorgänge länger andauern, kommunizieren Sie dies dem Benutzer. Beachten Sie die in der folgenden Tabelle aufgeführten Schwellenwerte. Siehe auch [Ressourcengrenzen und Leistungsoptimierung für Office Add-Ins](../../docs/develop/resource-limits-and-performance-optimization.md)


|**Interaktionsklasse**|**Ziel**|**Obergrenze**|**Menschlichen Wahrnehmung**|
|:-----|:-----|:-----|:-----|
|Sofort|<= 50 ms|100 ms|Ohne spürbaren Verzögerung.|
|Schnell|50-100 ms|200 ms|Minimal spürbare Verzögerung. Kein Feedback erforderlich.|
|Typisch|100-300 ms|500 ms|Schnell, aber zu langsam, um als schnell beschrieben zu werden. Kein Feedback erforderlich.|
|Reaktionsfähig|300-500 ms|1 Sekunde|Nicht schnell, aber dennoch reaktionsfähig. Kein Feedback erforderlich.|
|Fortlaufend.|> 500 ms|5 Sekunden|Mittlere Wartezeit, ist nicht mehr reaktionsfähig. Möglicherweise Feedback erforderlich.|
|Erfassung|> 500 ms|10 Sekunden|Lange, aber nicht lange genug, um etwas anderes zu tun. Möglicherweise Feedback erforderlich.|
|Erweitert|> 500 ms|> 10 Sekunden|Lange genug, um beim Warten etwas anderes zu tun. Möglicherweise Feedback erforderlich.|
|Lange ausgeführt|> 5 ms|> 1 minute|Benutzer tun in der Zwischenzeit mit Sicherheit etwas anderes tun.|
- Überwachen Sie den Dienststatus, und verwenden Sie Telemetrie, um erfolgreiche Benutzerausführungen zu überwachen.


## Vermarkten des Add-Ins



- Veröffentlichen Sie das Add-In im [Office Store ](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx), und [werben Sie dafür](http://msdn.microsoft.com/library/b19e21f8-76f5-44e1-9971-bef79cad4c71%28Office.15%29.aspx) auf Ihrer Website. Erstellen Sie eine [effektive Office Store-Liste](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

- Verwenden Sie kurze und prägnante beschreibende Add-In-Titel. Verwenden Sie nicht mehr als 128 Zeichen.

- Vermitteln Sie das Wertversprechen des Add-Ins im Titel und der Beschreibung. Verlassen Sie sich nicht auf die Marke.

- Vermitteln Sie das Wertversprechen des Add-Ins im Titel und der Beschreibung. Verlassen Sie sich nicht auf die Marke.

- Veröffentlichen Sie das Add-In im Office Store , und werben Sie dafür auf Ihrer Website.

