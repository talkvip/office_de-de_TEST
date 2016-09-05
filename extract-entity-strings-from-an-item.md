
# Extrahieren von Entitätszeichenfolgen aus einem Outlook-Element

In diesem Artikel wird die Erstellung des Outlook-Add-Ins  **Display entities** beschrieben, das Zeichenfolgeninstanzen von unterstützten bekannten Entitäten im Betreff und Textkörper des ausgewählten Outlook-Elements extrahiert. Dieses Element kann ein Termin, eine E-Mail oder eine Besprechungsanfrage, -antwort oder -stornierung sein. Die folgenden Entitäten werden unterstützt:

- Adresse: Eine Postadresse in den USA, die mindestens aus Teilen von Hausnummer, Straßenname, Stadt, Staat und Postleitzahl besteht
    
- Kontakt: Die Kontaktdaten einer Person im Kontext anderer Entitäten wie einer Adresse oder einem Unternehmensnamen
    
- E-Mail-Adresse: Eine SMTP-E-Mail-Adresse
    
- Besprechungsvorschlag: Ein Besprechungsvorschlag wie ein Verweis auf eine Veranstaltung. Beachten Sie, dass das Extrahieren von Besprechungsvorschlägen nur von Nachrichten, jedoch nicht von Terminen unterstützt wird.
    
- Telefonnummer: Eine nordamerikanische Telefonnummer
    
- Vorgangsvorschlag: Ein Vorgangsvorschlag, in der Regel ausgedrückt durch einen Ausdruck, für den eine Aktion ausgeführt werden kann
    
- URL
    
Die meisten dieser Entitäten basieren auf der Erkennung natürlicher Sprache. Zu diesem Zweck ist das Erlernen großer Datenmengen durch den Computer erforderlich. Diese Erkennung ist nicht deterministisch und manchmal vom Kontext im Outlook-Element abhängig. Outlook aktiviert das Entitäten-Add-In jedes Mal, wenn der Benutzer einen Termin, eine E-Mail oder eine Besprechungsanfrage, -antwort oder -stornierung zur Anzeige auswählt. Während der Initialisierung liest das Beispiel-Entitäten-Add-In alle Instanzen der unterstützten Entitäten des aktuellen Elements. 

Das Add-In stellt Schaltflächen bereit, mit denen der Benutzer einen Entitätstyp auswählen kann. Wenn der Benutzer eine Entität auswählt, zeigt das Add-In Instanzen der ausgewählten Entität im Add-in-Bereich an. In den folgenden Abschnitten werden die XML-Manifestdatei sowie HTML- und JavaScript-Dateien des Entitäten-Add-ins aufgeführt, und es wird der Code hervorgehoben, der die Extraktion der entsprechenden Entität unterstützt.

## XML-Manifestdatei


Das Entitäten-Add-In verfügt über die zwei Aktivierungsregeln, die durch einen logischen ODER-Operator verbunden werden: 


```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

Diese Regeln legen fest, dass Outlook dieses Add-in aktivieren soll, wenn das aktuell ausgewählte Element im Lesebereich oder Inspektor ein Termin oder eine Nachricht ist (einschließlich E-Mail, Besprechungsanfrage, -antwort oder -stornierung).

Im Folgenden ist das Manifest des Entitäten-Add-Ins angezeigt. Es verwendet Version 1.1 des Schemas für Office-Add-Ins-Manifeste.




```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
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
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## HTML-Implementierung


Die HTML-Datei des Entitäten-Add-Ins legt Schaltflächen fest, mit denen der Benutzer die einzelnen Entitätentypen auswählen kann, sowie eine weitere Schaltfläche, um die angezeigten Instanzen jeder Entität zu löschen. Es ist zudem eine JavaScript-Datei, default_entities.js, vorhanden, die im folgenden Abschnitt unter [JavaScript-Implementierung](#javascript-implementierung) beschrieben wird. Die JavaScript-Datei enthält die Ereignishandler für alle Schaltflächen.

Beachten Sie, dass alle Outlook-Add-Ins die Datei „office.js" enthalten müssen. Die folgende HTML-Datei enthält im CDN Version 1.1 der Datei „office.js". 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## Stylesheet


Das Entitäten-Add-In kann optional eine CSS-Datei, default_entities.css, verwenden, um das Layout der Ausgabe festzulegen. Die CSS-Datei wird im Anschluss aufgeführt.


```css
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## JavaScript-Implementierung


In den restlichen Abschnitten wird erläutert, wie dieses Beispiel (default_entities.js-Datei) bekannte Entitäten aus dem Betreff und Text der Nachricht oder des Termins extrahiert, das der Benutzer anzeigt. 


## Extrahieren von Entitäten bei der Initialisierung


Beim Auftreten des [Office.initiatlize](../../reference/shared/office.initialize.md)-Ereignisses ruft das Entitäten-Add-In die [getEntities](../../reference/outlook/Office.context.mailbox.item.md)-Methode des aktuellen Elements auf. Die  **getEntities**-Methode gibt die globale Variable  `_MyEntities` zurück, ein Array der Instanzen der unterstützten Entitäten. Das Folgende ist der entsprechende JavaScript-Code.


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## Extrahieren von Adressen


Wenn der Benutzer auf die Schaltfläche  **Adressen abrufen** klickt, ruft der Ereignishandler `myGetAddresses` von der [addresses](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von Adressen ab, falls Adressen extrahiert wurden. Jede extrahierte Adresse wird als Zeichenfolge in dem Array gespeichert.  `myGetAddresses` bildet eine lokale HTML-Zeichenfolge in .mdText`, um die Liste der extrahierten Adressen anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extrahieren von Kontaktinformationen


Wenn der Benutzer auf die Schaltfläche  **Kontaktinformationen abrufen** klickt, ruft der Ereignishandler `myGetContacts` von der [contacts](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von Kontakten zusammen mit ihren Informationen ab, wenn solche extrahiert wurden. Jeder extrahierte Kontakt wird als [Contact](../../reference/outlook/simple-types.md)-Objekt in dem Array gespeichert.  `myGetContacts` ruft weitere Daten zu jedem Kontakt ab. Beachten Sie, dass der Kontext entscheidend dafür ist, ob Outlook einen Kontakt aus einem Element extrahieren kann. Es muss am Ende der E-Mail eine Signatur vorhanden sein oder mindestens einige der folgenden Informationen in der Nähe des Kontakts:


- Die Zeichenfolge, die den Namen des Kontakts von der [Contact.personName](../../reference/outlook/simple-types.md)-Eigenschaft darstellt
    
- Die Zeichenfolge, die den mit dem Kontakt von der [Contact.businessName](../../reference/outlook/simple-types.md)-Eigenschaft verbundenen Unternehmensnamen darstellt
    
- Das mit dem Kontakt von der [Contact.phoneNumbers](../../reference/outlook/simple-types.md)-Eigenschaft verbundene Array von Telefonnummern. Jede Telefonnummer wird durch ein [PhoneNumber](../../reference/outlook/simple-types.md)-Objekt dargestellt
    
- Für jedes  **PhoneNumber**-Element im Telefonnummernarray die Zeichenfolge, die die Telefonnummer von der [PhoneNumber.phoneString](../../reference/outlook/simple-types.md)-Eigenschaft darstellt
    
- Das Array der mit dem Kontakt von der [Contact.urls](../../reference/outlook/simple-types.md)-Eigenschaft verbundenen URLs. Jede URL wird als Zeichenfolge in einem Arrayelement dargestellt.
    
- Das Array der mit dem Kontakt von der [Contact.emailAddresses](../../reference/outlook/simple-types.md)-Eigenschaft verbundenen E-Mail-Adressen. Jede E-Mail-Adresse wird als Zeichenfolge in einem Arrayelement dargestellt.
    
- Das Array der mit dem Kontakt von der [Contact.addresses](../../reference/outlook/simple-types.md)-Eigenschaft verbundenen Postadressen. Jede Postadresse wird als Zeichenfolge in einem Arrayelement dargestellt.
    
 `myGetContacts` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Daten für jeden Kontakt anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extrahieren von E-Mail-Adressen


Wenn der Benutzer auf die Schaltfläche  **E-Mail-Adressen abrufen** klickt, ruft der Ereignishandler `myGetEmailAddresses` von der [emailAddresses](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von SMTP-E-Mail-Adressen ab, wenn solche extrahiert wurden. Jede extrahierte E-Mail-Adresse wird als Zeichenfolge in dem Array gespeichert.  `myGetEmailAddresses` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Liste der extrahierten E-Mail-Adressen anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extrahieren von Besprechungsvorschlägen


Wenn der Benutzer auf die Schaltfläche  **Besprechungsvorschläge abrufen** klickt, ruft der Ereignishandler `myGetMeetingSuggestions` von der [meetingSuggestions](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von Besprechungsvorschlägen ab, wenn solche extrahiert wurden.


 >**Note**  Only messages but not appointments support the  **MeetingSuggestion** entity type.

Jeder extrahierte Besprechungsvorschlag wird im Array in einem [MeetingSuggestion](../../reference/outlook/simple-types.md)-Objekt gespeichert.  `myGetMeetingSuggestions` ruft weitere Daten zu jedem Besprechungsvorschlag ab:


- Die Zeichenfolge, die als Besprechungsvorschlag von der [MeetingSuggestion.meetingString](../../reference/outlook/simple-types.md)-Eigenschaft erkannt wurde
    
- Das Array der Besprechungsteilnehmer von der [MeetingSuggestion.attendees](../../reference/outlook/simple-types.md)-Eigenschaft. Jeder Teilnehmer wird durch ein [EmailUser](../../reference/outlook/simple-types.md)-Objekt dargestellt.
    
- Für jeden Teilnehmer den Namen von der [EmailUser.displayName](../../reference/outlook/simple-types.md)-Eigenschaft.
    
- Für jeden Teilnehmer die SMTP-Adresse von der [EmailUser.emailAddress](../../reference/outlook/simple-types.md)-Eigenschaft.
    
- Die den Ort des Besprechungsvorschlags darstellende Zeichenfolge von der [MeetingSuggestion.location](../../reference/outlook/simple-types.md)-Eigenschaft
    
- Die den Betreff des Besprechungsvorschlags darstellende Zeichenfolge von der [MeetingSuggestion.subject](../../reference/outlook/simple-types.md)-Eigenschaft
    
- Die die Anfangszeit des Besprechungsvorschlags darstellende Zeichenfolge von der [MeetingSuggestion.start](../../reference/outlook/simple-types.md)-Eigenschaft
    
- Die die Endzeit des Besprechungsvorschlags darstellende Zeichenfolge von der [MeetingSuggestion.end](../../reference/outlook/simple-types.md)-Eigenschaft
    
 `myGetMeetingSuggestions` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Daten für jeden Besprechungsvorschlag anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## Extrahieren von Telefonnummern


Wenn der Benutzer auf die Schaltfläche  **Telefonnummern abrufen** klickt, ruft der Ereignishandler `myGetPhoneNumbers` von der [phoneNumbers](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von Telefonnummern ab, wenn solche extrahiert wurden. Jede extrahierte Telefonnummer wird im Array in einem [PhoneNumber](../../reference/outlook/simple-types.md)-Objekt gespeichert.  `myGetPhoneNumbers` ruft weitere Daten zu jeder Telefonnummer ab:


- Die Zeichenfolge, die die Art der Telefonnummer darstellt, beispielsweise die Privattelefonnummer von der [PhoneNumber.type](../../reference/outlook/simple-types.md)-Eigenschaft
    
- Die Zeichenfolge, die die tatsächliche Telefonnummer von der [PhoneNumber.phoneString](../../reference/outlook/simple-types.md)-Eigenschaft darstellt
    
- Die Zeichenfolge, die ursprünglich als Telefonnummer von der [PhoneNumber.originalPhoneString](../../reference/outlook/simple-types.md)-Eigenschaft erkannt wurde
    
 `myGetPhoneNumbers` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Daten für jede Telefonnummer anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extrahieren von Vorgangsvorschlägen


Wenn der Benutzer auf die Schaltfläche  **Vorgangsvorschläge abrufen** klickt, ruft der Ereignishandler `myGetTaskSuggestions` von der [taskSuggestions](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von Vorgangsvorschlägen ab, wenn solche extrahiert wurden. Jeder extrahierte Vorgangsvorschlag wird im Array in einem [TaskSuggestion](../../reference/outlook/simple-types.md)-Objekt gespeichert.  `myGetTaskSuggestions` ruft weitere Daten zu jedem Vorgangsvorschlag ab:


- Die Zeichenfolge, die ursprünglich als Vorgangsvorschlag von der [TaskSuggestion.taskString](../../reference/outlook/simple-types.md)-Eigenschaft erkannt wurde
    
- Das Array der dem Vorgang Zugewiesenen von der [TaskSuggestion.assignees](../../reference/outlook/simple-types.md)-Eigenschaft. Jeder Zugewiesene wird durch ein [EmailUser](../../reference/outlook/simple-types.md)-Objekt dargestellt.
    
- Für jeden Zugewiesenen den Namen von der [EmailUser.displayName](../../reference/outlook/simple-types.md)-Eigenschaft.
    
- Für jeden Zugewiesenen die SMTP-Adresse von der [EmailUser.emailAddress](../../reference/outlook/simple-types.md)-Eigenschaft.
    
 `myGetTaskSuggestions` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Daten für jeden Vorgangsvorschlag anzuzeigen. Das Folgende ist der entsprechende JavaScript-Code.




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Extrahieren von URLs


Wenn der Benutzer auf die Schaltfläche  **URLs abrufen** klickt, ruft der Ereignishandler `myGetUrls` von der [urls](../../reference/outlook/simple-types.md)-Eigenschaft des  `_MyEntities`-Objekts ein Array von URLs ab, wenn solche extrahiert wurden. Jede extrahierte URL wird als Zeichenfolge in dem Array gespeichert.  `myGetUrls` bildet eine lokale HTML-Zeichenfolge in `htmlText`, um die Liste der extrahierten URLs anzuzeigen.


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Löschen von angezeigten Entitätszeichenfolgen


Schließlich legt das Entitäten-Add-In einen  `myClearEntitiesBox`-Ereignishandler fest, der die angezeigten Zeichenfolgen löscht. Das Folgende ist der entsprechende Code.


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## JavaScript-Auflistung


Das Folgende ist die vollständige Auflistung der JavaScript-Implementierung.


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## Weitere Ressourcen



- [Erstellen von Outlook-Add-Ins für Leseformulare](../outlook/read-scenario.md)
    
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [item.getEntities-Methode](../../reference/outlook/Office.context.mailbox.item.md)
    
