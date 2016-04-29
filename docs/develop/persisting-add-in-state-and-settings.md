
# Beibehalten des Zustands und der Einstellungen von Add-ins
In diesem Thema wird beschrieben, welche Optionen für den dauerhaften Add-in-Status und Einstellungsdaten für Inhalt, Aufgabenbereich und Outlook-Add-Ins zur Verfügung stehen. Erfahren Sie mehr über die gemeinsamen Funktionen und Unterschiede zwischen den Stausi der  **Settings**-,  **RoamingSettings**- und  **CustomProperties**-Objekten der JavaScript-API für Office für den dauerhaften Add-In-Status, und sehen Sie sich Beispiele an, wie sie verwendet werden können.

 _**Gilt für:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

Office-Add-Ins sind grundlegende Web-Anwendungen, die in einer zustandslosen Umgebung eines Browsersteuerelements ausgeführt werden. Dadurch kann es für Ihr Add-in erforderlich sein, dass Daten gespeichert werden müssen, damit die Kontinuität bestimmter Vorgänge oder Funktionen sitzungsübergreifend gewährleistet werden können. Wenn Ihr Add-in beispielsweise benutzerdefinierte Einstellungen oder andere Werte aufweist, die gespeichert und bei der nächsten Initialisierung neu geladen werden müssen, wie z. B. eine vom Benutzer bevorzugte Ansicht oder ein Standardspeicherort.

Dazu können Sie:


- Verwenden Sie die Mitglieder der JavaScript-API für Office, die Daten als Namen-/Wertepaare in einer Eigenschaftensammlung an einem Add-in-abhängigen Speicherort ablegt.
    
- Verwenden Sie Verfahren des zugrunde liegenen Browsersteuerelements: Browsercookies oder HTML5 Webspeicher ([localStorage](http://msdn.microsoft.com/de-de/library/cc848902%28v=vs.85%29.aspx) or [sessionStorage](http://msdn.microsoft.com/de-de/library/cc197020%28v=vs.85%29.aspx)).
    
Der Schwerpunkt dieses Artikelts liegt auf der Verwendung der JavaScript-API für Office, um einen beständigen Add-in-Status zu erreichen. Beispiele für die Verwendung von Browsercookies und Webspeicher finden Sie unter [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## Mit der JavaScript-API für Office gespeicherter Add-in-Status und Einstellungen


Die JavaScript-API für Office stellt die [Settings](http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)-, [RoamingSettings](../reference/outlook/RoamingSettings.html%28Office.15%29.md)- und [CustomProperties](https://dev.outlook.com/reference/add-ins/CustomProperties.md%28Office.15%29.md)-Objekte zum Speichern des Add-In-Status sitzungsübergreifend bereit, wie in der folgenden Tabelle beschrieben. In allen Fällen sind die gespeicherten Einstellungswerte nur der [ID](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) des Add-Ins zugeordnet, das sie erstellt hat.



|**Objekt**|**Unterstützung für Add-in-Typen**|**Speicherort**|**Officehost-Unterstützung**|
|:-----|:-----|:-----|:-----|
|[Settings](http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)|Speichern von App-Status und Einstellungen pro Dokument für Inhalts- und Aufgabenbereich-Apps|Das Dokument, Arbeitsblatt oder die Präsentation, mit dem bzw. der das Add-in arbeitet.Inhalts- und Aufgabenbereich-Add-in-Einstellungen sind für das Add-in verfügbar, das sie aus dem Dokument erstellt hat, in dem sie gespeichert werden. **Wichtig:** Speichern Sie keine Kennwörter oder andere personenbezogenen Daten (PII) mit dem **Settings**-Objekt. Die gespeicherten Daten sind für Endbenutzer nicht sichtbar, sie werden jedoch as Teil des Dokuments gespeichert, auf das beim Lesen des Dokumentdateiformats direkt zugegriffen wird. Sie sollten die Verwendung von personenbezogenen Daten durch das Add-in einschränken und alle für das Add-in erforderlichen personenbezogenen Daten nur auf dem Server speichern, auf dem das Add-in als benutzergeschützte Ressource gehostet wird.|Word, Excel oder PowerPoint **Hinweis:** Aufgabenbereich-Add-ins für Project 2013 unterstützen die **Settings**-API zum Speichern des Add-in-Status oder der Einstellungen nicht. Für in Project ausgeführte Add-ins (sowie andere Office-Hostanwendungen) können Sie Verfahren wie die Verwendung von Browsercookies oder Webspeicherung nutzen. Weitere Informationen über diese Verfahren finden Sie unter [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../reference/outlook/RoamingSettings.md%28Office.15%29.md)|Outlook|Das Postfach des Benutzers auf dem Exchange Server, auf dem das Add-in installiert ist.Da diese Einstellungen im Postfach des Benutzers gespeichert werden, sind sie für Benutzer unterwegs und für das Add-in verfügbar, sobald dieses im Kontext einer beliebigen unterstützten Clienthostanwendung oder eines Browsers ausgeführt wird, mit der bzw. dem auf das Postfach des Benutzers zugegriffen wird. Die Outlook-Add-In-Roamingeinstellungen stehen nur dem Add-In zur Verfügung, das sie erstellt hat und nur von dem Postfach, in dem das Add-In installiert ist.|Outlook|
|[CustomProperties](../reference/outlook/CustomProperties.md%28Office.15%29.md)|Outlook|Das Nachrichten-, Termin- oder Besprechungsanfragenelement, mit dem das Add-in arbeitet. Benutzerdefinierte Outlook-Add-In-Elementeigenschaften stehen nur dem Add-In zur Verfügung, das sie erstellt hat und nur von dem Element aus, in dem sie gespeichert sind.|Outlook|

## Einstellungsdaten werden im Speicher zur Laufzeit verwaltet.


Die Daten im Eigenschaftenbehälter, auf die unter Verwendung des  **Settings**-,  **CustomProperties**- oder  **RoamingSettings**-Objekts zugegriffen wird, werden als serialisiertes JSON-Objekt (JavaScript Object Notation) gespeichert. Im folgenden Beispiel wird die Struktur des Eigenschaftenbehälters veranschaulicht, der drei definierte  **string**-Werte mit den Namen  `app_setting_name_0`,  `app_setting_name_1` und `app_setting_name_2` enthält.

Dieser Beispiel der Eigenschaftssammlungsstruktur enthält drei definierte  **string** Werte namens `firstName`,  `location` und `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

Der Einstellungseigenschaftenbehälter kann bei der Initialisierung des Add-ins oder zu einem beliebigen späteren Zeitpunkt während der aktuellen Sitzung des Add-ins geladen werden. Während der Sitzung wird der Einstellungseigenschaftenbehälter vollständig im Speicher verwaltet. Dabei werden die Methoden  **get**,  **set** und **remove** des Objekts verwendet, das der Art von Einstellungen entspricht, die Sie erstellen und mit denen Sie arbeiten ( **Settings**,  **CustomProperties** oder **RoamingSettings**). 


 >**Wichtig**  Um alle während der aktuellen Add-in-Sitzung vorgenommenen Ergänzungen, Aktualisierungen oder Löschungen am Speicherort beizubehalten, müssen Sie die  **saveAsync**-Methode des verwendeten Objekts für diese Art von Einstellungen aufrufen. Die Methoden  **get**,  **set** und **remove** werden nur auf die speicherinterne Kopie des Einstellungseigenschaftenbehälters angewendet. Wenn Ihr Add-in ohne Aufruf von **saveAsync** geschlossen wird, gehen alle Änderungen verloren, die während der Sitzung an den Einstellungen vorgenommen wurden.


## So speichern Sie Add-in-Status und -Einstellungen für Inhalts- und Aufgabenbereich-Add-ins


Verwenden Sie zum Speichern des Status oder der benutzerdefinierten Einstellungen eines Inhalts- oder Aufgabenbereich-Add-ins das [Settings](http://msdn.microsoft.com/de-de/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)-Objekt und seine Elemente. Der mithilfe der Methoden des  **Settings**-Objekts erstellte Eigenschaftenbehälter wird pro Add-in und pro Dokument gespeichert. Dies bedeutet, dass dieser nur für die Instanz des Inhalts- oder Aufgabenbereich-Add-ins verfügbar ist, die diesen erstellt hat und ausschließlich von dem Dokument, in dem er gespeichert wurde.

Das  **Settings**-Objekt wird automatisch als Bestandteil des [Document](http://msdn.microsoft.com/de-de/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)-Objekts geladen, und es ist verfügbar, wenn das Aufgabenbereich- oder Inhalts-Add-in aktiviert wird. Nach der Instanziierung des  **Document**-Objekts können Sie auf das Objekt  **Settings** mit der Eigenschaft [settings](http://msdn.microsoft.com/de-de/library/77ba7daf-419f-44b6-8747-7fd5618b7053%28Office.15%29.aspx) des Objekts **Document** zugreifen. Während der Dauer der Sitzung können Sie einfach die Methoden **Settings.get**,  **Settings.set** und **Settings.remove** verwenden, um dauerhafte Einstellungen und den Add-in-Status aus der speicherinternen Kopie des Eigenschaftenbehälters zu lesen, zu schreiben oder zu entfernen.

Da die Festlegen- und Entfernen-Methoden nur gegen die sich im Speicher befindliche Kopie der Einstellungseigenschaftensammlung agieren, wird das Add-in zum Speichern neuer oder geänderter Einstellungen für das Dokument der [Settings.saveAsync](http://msdn.microsoft.com/de-de/library/7147c221-937c-477c-98a6-f59d6200c27b%28Office.15%29.aspx)-Methode zugeordnet.


### Erstellen und Aktualisieren eines Einstellungswerts

Das folgende Codebeispiel veranschaulicht, wie die [Settings.set](http://msdn.microsoft.com/de-de/library/4e2c9758-953e-41e8-aca6-d8daf764a584%28Office.15%29.aspx)-Methode verwendet werden kann, um eine Einstellung namens  `'themeColor'` mit einem Wert `'green'` zu erstellen. Beim ersten Parameter der set-Methode, _name_ (ID), wird zwischen Groß- und Kleinschreibung der festzulegenden oder zu erstellenden Einstellung unterschieden. Der zweite Parameter ist der _value_ der Einstellung.


```
Office.context.document.settings.set('themeColor', 'green');
```

 Die Einstellung wird mit dem angegebenen Namen erstellt, falls sie nicht bereits vorhanden ist oder ihr Wert wird aktualisiert, fall sie vorhanden ist. Verwenden Sie die **Settings.saveAsync**-Methode zum Speichern der neuen oder aktualisierten Einstellungen für das Dokument.


### Abrufen des Wertes einer Einstellung

Das folgende Beispiel veranschaulicht die Verwendung der [Settings.get](http://msdn.microsoft.com/de-de/library/aeac06dd-994e-4235-b208-1bd117395296%28Office.15%29.aspx)-Methode, um den Wert einer Einstellung namens "themeColor" abzurufen. Der einzige Parameter der  **get**-Methode ist der  _name_ der Einstellung, für den Groß-/Kleinschreibung unterschieden wird.


```
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 Der einzige Parameter der **get**-Methode ist der  _name_ der Einstellung, in dem Groß- und Kleinschreibung unterschieden wird. Die **get**-Methode gibt den Wert zurück, der zuvor für den übergebenen Einstellungsnamen gespeichert wurde. Wenn die Einstellung nicht vorhanden ist, gibt die Methode  **null** zurück.


### Entfernen einer Einstellung

Das folgende Beispiel veranschaulicht die Verwendung der [Settings.remove](http://msdn.microsoft.com/de-de/library/a92446bf-de65-45bd-8412-36ea8e77c5a2%28Office.15%29.aspx)-Methode, um eine Einstellung namens "themeColor" zu entfernen. Der einzige Parameter der  **remove**-Methode ist der  _name_ der Einstellung, für den Groß-/Kleinschreibung unterschieden wird.


```
Office.context.document.settings.remove('themeColor');
```

Es passiert nichts, wenn die Einstellung nicht vorhanden ist. Verwenden Sie die  **Settings.saveAsync**-Methode, um die Entfernung der Einstellung aus dem Dokument zu speichern.


### Speichern Ihrer Einstellungen

Um die Hinzufügungen, Änderungen und Entfernungen an Ihrem Add-in in der Speicherkopie der Einstellungs-Eigenschaftssammlung zu speichern, müssen Sie die [Settings.saveAsync](http://msdn.microsoft.com/de-de/library/7147c221-937c-477c-98a6-f59d6200c27b%28Office.15%29.aspx)-Methode aufrufen, damit sie im Dokument gespeichert werden. Der einzige Parameter der  **saveAsync**-Methode ist der  _callback_, der eine Rückruffunktion mit einem Parameter darstellt. 


```
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Die anonyme Funktion wurde an die Methode  **saveAsync** übergeben, als der Parameter _callback_ beim Abschluss des Vorgangs ausgeführt wurde. Die Rückruffunktion wird mit einem einzigen Parameter aufgerufen, _asyncResult_, der den Zugriff auf ein  **AsyncResult**-Objekt ermöglicht, das den Status des Vorgangs enthält. Im Beispiel überprüft die Funktion die Eigenschaft  **AsyncResult.status**,um zu sehen, ob der Speichervorgang erfolgreich oder fehlerhaft war, und zeigt dann das Ergebnis auf der Seite des Add-ins an.


## So speichern Sie Einstellungen im Postfach des Benutzers für Outlook-Add-Ins als Roamingeinstellungen


Ein Outlook-Add-In kann das [RoamingSettings](../reference/outlook/RoamingSettings.md%28Office.15%29.md)-Objekt verwenden, um für das Postfach des Benutzers spezifische Daten zu speichern. Auf diese Daten kann nur durch dieses Outlook-Add-In zugegriffen werden. Die Daten werden auf dem Exchange-Server für dieses Postfach gespeichert, und auf diese kann zugegriffen werden, wenn sich der Benutzer bei seinem Konto anmeldet und das Outlook-Add-In ausführt.


### Laden von Roamingeinstellungen


Ein Outlook-Add-In lädt Roamingeinstellungen in der Regel im [Office.initialize](http://msdn.microsoft.com/en-us/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignishandler. Im folgenden JavaScript-Codebeispiel wird das Laden vorhandener Roamingeinstellungen veranschaulicht.


```
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### Erstellen oder Zuweisen einer Roamingeinstellung


Ausgehend vom vorherigen Beispiel veranschaulicht die folgende JavaScript-Funktion  `setAppSetting`, wie die [RoamingSettings.set](../reference/outlook/RoamingSettings.html%28Office.15%29.md)-Methode verwendet wird, um eine Einstellung namens  `cookie` mit dem heutigen Datum festzulegen. Anschließend speichert sie alle Roamingeinstellungen mit der [RoamingSettings.saveAsync](https://dev.outlook.com/reference/add-ins/RoamingSettings.md%28Office.15%29.md)-Methode wieder auf dem Exchange-Server. 


```
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

Die  **saveAsync**-Methode speichert Roamingeinstellungen asynchron und akzeptiert eine optionale Rückruffunktion. Dieses Codebeispiel gibt eine Rückruffunktion namens  `saveMyAppSettingsCallback` an die **saveAsync**-Methode weiter. Wenn der asynchrone Aufruf zurückgegeben wird, stellt der  _asyncResult_ Parameter der `saveMyAppSettingsCallback`-Funktion einen Zugang zum [AsyncResult](../reference/outlook/simple-types.md%28Office.15%29.md)-Objekt bereit, das Sie verwenden können, um den Erfolg oder Misserfolg des Vorgangs mit der  **AsyncResult.status**-Eigenschaft zu ermitteln.


### Entfernen einer Roamingeinstellung


Ebenfalls ausgehend von den vorherigen Beispielen veranschaulicht die folgende  `removeAppSetting`-Funktion, wie die [RoamingSettings.remove](../reference/outlook/RoamingSettings.md%28Office.15%29.md)-Methode verwendet wird, um die  `cookie`-Einstellung zu entfernen und alle Roamingeinstellungen wieder auf dem Exchange-Server zu speichern.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## So speichern Sie Einstellungen pro Element für Outlook-Add-Ins als benutzerdefinierte Eigenschaften


Benutzerdefinierte Eigenschaften ermöglichen das Speichern von Informationen zu einem Element. Wenn Sie beispielsweise ein Outlook-Add-In erstellen, das einen Termin aus einer Besprechungsanfrage in einer Nachricht erstellt, können Sie benutzerdefinierte Eigenschaften verwenden, um die Tatsache zu speichern, dass der Termin erstellt wurde. Auf diese Weise wird sichergestellt, dass das Outlook-Add-In bei einem erneuten Öffnen der Nachricht nicht erneut vorschlägt, einen Termin zu erstellen.

Bevor Sie benutzerdefinierte Eigenschaften für eine bestimmte Nachricht, einen bestimmten Termin oder eine bestimmte Besprechungsanfrage verwenden können, müssen Sie die Eigenschaften in den Speicher laden, indem Sie die [loadCustomPropertiesAsync](../reference/outlook/Office.context.mailbox.item.html%28Office.15%29.md)-Methode des  **Item**-Objekts laden. Wenn für das aktuelle Element bereits benutzerdefinierte Eigenschaften festgelegt wurden, werden sie zu diesem Zeitpunkt vom Exchange-Server geladen. Nachdem Sie die Eigenschaften geladen haben, können Sie die [set](https://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)- und die [get](https://dev.outlook.com/reference/add-ins/RoamingSettings.html%28Office.15%29.md)-Methode des  **CustomProperties**-Objekts verwenden, um Eigenschaften im Speicher hinzuzufügen, zu aktualisieren und abzurufen. Zum Speichern der an den benutzdefinierten Eigenschaften des Elements vorgenommenen Änderungen müssen Sie die [saveAsync](https://dev.outlook.com/reference/add-ins/CustomProperties.md%28Office.15%29.md)-Methode verwenden, um die Änderungen an dem Element auf dem Exchange-Server zu speichern.


### Beispiel für benutzerdefinierte Eigenschaften

Im folgenden Beispiel wird eine vereinfachte Gruppe von Methoden für einOutlook-Add-In dargestellt, das benutzerdefinierte Eigenschaften verwendet. Sie können dieses Beispiel als Ausgangspunkt für Ihr Outlook-Add-In mit benutzerdefinierten Eigenschaften verwenden. 

Ein Outlook-Add-In, das diesen Code verwendet, ruft alle benutzerdefinierten Eigenschaften durch Aufruf der  **get**-Methode auf der  `_customProps`-Variable ab, wie im folgenden Beispiel dargestellt.




```
var property = _customProps.get("propertyName");
```

Dieses Beispiel beinhaltet die folgenden Funktionen:



|**Funktionsname**|**Beschreibung**|
|:-----|:-----|
| `Office.initialize`|Initialisiert das Add-in und lädt den benutzerdefinierten Eigenschaftenbehälter vom Exchange-Server.|
| `customPropsCallback`|**customPropsCallback**: Ruft den vom Exchange-Server zurückgegebenen benutzerdefinierten Eigenschaftenbehälter ab und speichert ihn zur späteren Verwendung|
| `updateProperty`|**updateProperty**: Legt eine bestimmte Eigenschaft fest oder aktualisiert diese und speichert die Änderungen anschließend auf dem Exchange-Server|
| `removeProperty`|**removeProperty**: Entfernt eine bestimmte Eigenschaft aus dem Eigenschaftenbehälter und speichert die Entfernung anschließend auf dem Exchange-Server|
| `saveCallback`|Rückruf für Aufrufe der  **saveAsync**-Methode in den  `updateProperty`- und  `removeProperty`-Funktionen.|



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## Zusätzliche Ressourcen



- [Grundlegendes zur JavaScript-API für Office](01180dae-ca45-40c8-b3dd-fd2a85651c0c.md)
    
- [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
