
# XML-Manifest für Office-Add-Ins
Erstellen oder bearbeiten Sie eine XML-Manifestdatei für eine Office-Add-In.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Die XML-Manifestdatei einer Office-Add-In ermöglicht Ihnen die deklarative Beschreibung, wie Ihr Add-in aktiviert werden soll, wenn ein Endbenutzer dieses mit Office-Dokumenten und -Anwendungen installiert und verwendet. Die Datei [offappmanifest.xsd](http://msdn.microsoft.com/de-de/library/d5f72bff-3446-c64f-02ca-ab10b5648789%28Office.15%29.aspx) definiert ein XML-Schema, das allen unterstützten Office-Anwendungen gemeinsam ist (sowohl Rich Client-Anwendungen als auch deren entsprechenden Web App-Webclients).

 >**Hinweis**  Ihre Office-Add-In muss Manifestschema Version 1.1 verwenden. Informationen zum Aktualisieren Ihrer Add-In für die Verwendung von Manifest 1.1 finden Sie unter [Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Eine auf diesem Schema basierende XML-Manifestdatei ermöglicht einer Office-Add-In Folgendes:

- Beschreibung von sich selbst durch die Bereitstellung einer ID, einer Version, einer Beschreibung, eines Anzeigenamens und eines Standardgebietsschemas
    
- Angabe des Speicherorts der HTML-Datei, die die Benutzeroberfläche der Office-Add-In bereitstellt
    
- Angabe der angeforderten Standardabmessungen für Inhalts-Office-Add-Ins und der angeforderten Höhe für Outlook-Office-Add-Ins
    
- Deklaration der von der Office-Add-In benötigten Berechtigungen wie Lesen oder Schreiben in dem Dokument
    
- Festlegung, wie diese verwendet und angezeigt werden sollen, beispielsweise als Inhalt (im Dokument), in einem Aufgabenbereich oder kontextbezogen mit einer Nachricht, einem Termin oder einer Besprechungsanfrage
    
- Für Outlook-Office-Add-Ins: Definition der Regeln, die den Kontext angeben, in dem sie aktiviert werden und mit dem Nachrichten-, Termin- oder Besprechungsanfrageelementen interagieren
    
Beispiele von XML-Manifestdateien v1.1 finden Sie unter [Manifest v1.1 XML-Dateibeispiele](#manifest-v1.1-xml-dateibeispiele).

## Erforderliche Elemente


In der folgenden Tabelle werden die Elemente angegeben, die für die drei Arten der Office-Add-Ins und das Wörterbuch-Aufgabenbereich-Add-in erforderlich sind.


 >**Wichtig**  Für an den Office Store gesendete Add-ins müssen alle Speicherorte des Add-ins wie etwa die im  **SourceLocation**-Element angegebenen Speicherorte für die Quelldatei mit SSL gesichert sein (HTTPS). Weitere Informationen finden Sie unter [Welche häufig auftretenden Fehler beim Senden gilt es zu vermeiden?](http://msdn.microsoft.com/library/0ceb385c-a608-40cc-8314-78e39d6c75d0%28Office.15%29.aspx#bk_q2)


**Erforderliche Elemente nach Office-Add-in-Typ**


|**Element**|**Inhalt**|**Aufgabenbereich**|**Wörterbuch-Aufgabenbereich**|**Outlook**|
|:-----|:-----|:-----|:-----|:-----|
|[OfficeApp](http://msdn.microsoft.com/de-de/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx)|X|X|X|X|
|[Id](http://msdn.microsoft.com/de-de/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)|X|X|X|X|
|[Version](http://msdn.microsoft.com/de-de/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx)|X|X|X|X|
|[ProviderName](http://msdn.microsoft.com/de-de/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx)|X|X|X|X|
|[DefaultLocale](http://msdn.microsoft.com/de-de/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx)|X|X|X|X|
|[DisplayName](http://msdn.microsoft.com/de-de/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx)|X|X|X|X|
|[Description](http://msdn.microsoft.com/de-de/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx)|X|X|X|X|
|[IconUrl](http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)|X|X|X|X|
|[HighResolutionIconUrl](http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx)||||X|
|[DefaultSettings (ContentApp)](http://msdn.microsoft.com/de-de/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx)[DefaultSettings (TaskPaneApp)](http://msdn.microsoft.com/de-de/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx)|X|X|X||
|[SourceLocation (ContentApp)](http://msdn.microsoft.com/de-de/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx)[SourceLocation (TaskPaneApp)](http://msdn.microsoft.com/de-de/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)|X|X|X||
|[DesktopSettings](http://msdn.microsoft.com/de-de/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx)||||X|
|[SourceLocation (MailApp)](http://msdn.microsoft.com/de-de/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx)||||X|
|[Permissions (ContentApp)](http://msdn.microsoft.com/de-de/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx)[Permissions (TaskPaneApp)](http://msdn.microsoft.com/de-de/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx)[Permissions (MailApp)](http://msdn.microsoft.com/de-de/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)|X|X|X|X|
|[Rule (RuleCollection)](http://msdn.microsoft.com/de-de/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx)[Rule (MailApp)](http://msdn.microsoft.com/de-de/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)||||X|
|[Dictionary](http://msdn.microsoft.com/de-de/library/f78898f4-059e-d5dc-5eab-1f6b92214068%28Office.15%29.aspx)|||X||
|[*Requirements (MailApp)](http://msdn.microsoft.com/de-de/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx)||||X|
|[*Set](http://msdn.microsoft.com/de-de/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx)[**Sets (MailAppRequirements)](http://msdn.microsoft.com/de-de/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx)||||X|
|[*Form](http://msdn.microsoft.com/de-de/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx)[**FormSettings](http://msdn.microsoft.com/de-de/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx)||||X|
|[*Sets (Requirements)](http://msdn.microsoft.com/de-de/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx)||||X|
|[*Hosts](http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx)||||X|
* In Office-Add-In Manifest-Schemaversion 1.1 hinzugefügt.


## Manifest v1.1 XML-Dateibeispiele


Der folgende Abschnitt enthält Beispiele von XML-Manifestdateien v1.1 für Inhalts-, Aufgabenbereich-, Outlook- und Wörterbuch-Office-Add-Ins.

Falls Sie Visual Studio verwenden, um Ihre Office-Add-In zu entwickeln, können Sie den Visual Studio Manifest-Designer verwenden, um die Office-Add-In-Manifesteinstellungen zu ändern, statt manuell das zugrunde liegende XML-Markup zu ändern. Standardmäßig wird beim Öffnen einer Office-Add-In-Manifestdatei in Visual Studio der Manifest-Designer geöffnet. Der Designer organisiert die Felder in der Manifestdatei, sodass man sie leichter finden kann. Einige Felder besitzen Dropdown-Listenfelder, die bestimmte gültige Feldwerte enthalten, wodurch Dateneingabefehler verringert werden.


### Beispiel-Manifestdatei für Inhalts-Office-Add-In v1.1



```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth> 
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```


### Beispiel-Manifestdatei für Aufgabenbereich-Office-Add-In v1.1



```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="TaskPaneApp">
  <Id>412ce350-4161-4ad0-a5f5-0ec9d2cd3570</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample task pane add-in">
    <Override Locale="en-US" Value="Project add-in"/> 
  </DisplayName>
  <Description DefaultValue="Describe the features of this app.">
    <Override Locale="en-US" Value="Adds project management information to documents"/>
  </Description>
  <IconUrl DefaultValue="https://contoso.com.sa/ProjectApp/Topgunas-SA.png">
    <Override Locale="en-US" Value="https://contoso.com/ProjectApp/Topgunen-US.png"/>
  </IconUrl>
  <AppDomains>
    <AppDomain>http://www.projectlogin.com</AppDomain>
    <AppDomain>http://m.projectlogin.com</AppDomain>
    <AppDomain>http://www.projectlogin.com.sa</AppDomain>
    <AppDomain>http://m.projectlogin.com.sa</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name = "Document"/>
    <Host Name = "Workbook"/>
    <Host Name = "Presentation"/>
    <Host Name = "Project"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com.sa/ProjectApp/ProjectAppar_SA.html">
      <Override Locale="en-US" Value="https://contoso.com/ProjectApp/ProjectAppen-US.html"/>
    </SourceLocation>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```


### Beispiel-Manifestdatei für Outlook-Office-Add-In v1.1



```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube référencées dans vos courriers électronique
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or"> 
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch" 
        PropertyName="BodyAsPlaintext" RegExName="VideoURL" 
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
    </Rule> 
  </Rule>
</OfficeApp>

```


### Beispielmanifestdatei für Wörterbuch-Office-Add-In v1.1



```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="TaskPaneApp">
  <Id>6f28f837-8326-4231-a033-ea1d80ff9fb3</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="https://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <Hosts>
    <Host Name = "Document"/>
    <Host Name = "Workbook"/>
    <Host Name = "Presentation"/>
    <Host Name = "Project"/>
  </Hosts>
  <DefaultSettings>
    <!--Replace the value specified for DefaultValue below with the URL for your dictionary-->
    <SourceLocation DefaultValue="https://contoso.com/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of dialects your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. This is for different dialects of the same language. Please do NOT put more than one language (e.g. Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML webservice (which we use to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="https://contoso.com/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line - e.g. this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="https://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```


## Überprüfen des Office-Add-Ins-Manifests


Um sicherzustellen, dass die Manifestdatei, die Ihre Office-Add-In beschreibt, korrekt und vollständig ist, überprüfen Sie diese anhand der XSD-Datei (XML Schema Definition). Sie können ein XML-Schemaüberprüfungstool oder Visual Studio zum Überprüfen des Manifests verwenden. Sie können das [Office-App-Kompatibilitätskit](https://www.microsoft.com/en-us/download/details.aspx?id=46831) und führen Sie es bei Ihrem Add-In aus.

Informationen zum Prüfen einer Manifestdatei auf Gültigkeit für ein Schema finden Sie unter [XML Schemaüberprüfungstool (XSD)](http://stackoverflow.com/questions/124865/xml-schema-xsd-validation-tool).


 >**Hinweis**  Die Überprüfung Ihres Manifest anhand der Datei "offappmanifest.xsd" ist veraltet. Informationen zum Aktualisieren Ihres Manifests auf die Schemaversion 1.1 finden Sie unter [Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).


## Angeben von Domänen, die im Add-in-Fenster geöffnet werden sollen


Standardmäßig wird, wenn Ihr Add-in versucht, zu einer URL zu navigieren, die sich in einer anderen Domäne befindet als der, in der die Startseite gehostet wird (im [SourceLocation](http://msdn.microsoft.com/de-de/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx)-Element der Manifestdatei angegeben), diese URL in einem neuen Browserfenster außerhalb des Add-in-Bereichs der Office-Hostanwendung geöffnet. Durch dieses Standardverhalten wird der Benutzer vor einer unerwarteten Seitennavigation innerhalb des Add-in-Bereichs durch eingebettete  **iframe**-Elemente geschützt. 

Um dieses Verhalten außer Kraft zu setzen, geben Sie in der Liste der Domänen im [AppDomains](http://msdn.microsoft.com/de-de/library/13cf867d-9b24-786f-0687-6bcdc954628e%28Office.15%29.aspx)-Element der Manifestdateialle Domänen an, die im Add-in-Fenster geöffnet werden sollen. Versucht das Add-in, zu einer URL zu navigieren, die sich nicht auf der Liste befindet, wird die URL in einem neuen Browserfenster (außerhalb des Add-in-Bereichs) geöffnet.

Im folgenden XML-Manifestbeispiel wird die Haupt-Add-in-Seite in der Domäne  `https://www.contoso.com` gehostet, wie im Element **SourceLocation** angegeben. Außerdem wird die Domäne `https://www.northwindtraders.com` in einem [AppDomain](http://msdn.microsoft.com/de-de/library/2a0353ec-5e09-6fbf-1636-4bb5dcebb9bf%28Office.15%29.aspx)-Element in der  **AppDomains**-Elementliste angegeben. Wenn das Add-in zu einer Seite in der Domäne „www.northwindtraders.com" navigiert, wird diese Seite im Add-in-Bereich geöffnet.




```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```


## Zusätzliche Ressourcen



- [Angeben von Office-Hosts und API-Anforderungen](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [Erstellen von Add-In-Befehlen in Ihrem Manifest für Excel, Word und PowerPoint (Vorschau)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md)
    
- [Lokalisierung für Office-Add-Ins](../../develop/localization.md)
    
- [Schemareferenz für Office-Add-ins-Manifeste (v1. 1) ](http://msdn.microsoft.com/de-de/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92%28Office.15%29.aspx)
    
