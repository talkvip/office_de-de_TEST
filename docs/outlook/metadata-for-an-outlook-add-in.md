
# Abrufen und Festlegen von Add-In-Metadaten für ein Outlook-Add-In
Erhalten Sie Informationen zum Laden und Speichern benutzerdefinierter Daten, auf die nur Ihr Outlook-Add-In zugreifen kann – über Roamingeinstellungen oder benutzerdefinierte Eigenschaften.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Es gibt zwei Methoden zum Verwalten von benutzerdefinierter Daten in Ihrem Outlook-Add-In:

- Roamingeinstellungen, mit denen benutzerdefinierte Daten für das Postfach des Benutzers verwaltet werden.
    
- Benutzerdefinierte Eigenschaften, mit denen benutzerdefinierte Daten für ein Element im Postfach des Benutzers verwaltet werden.
    
Beide gewähren Zugriff auf benutzerdefinierte Daten, auf die nur über Ihr Outlook-Add-In zugegriffen werden kann, jede Methode speichert die Daten jedoch getrennt voneinander, d. h. Benutzerdefinierte Eigenschaften können nicht auf über Roamingeinstellungen gespeicherten Daten zugreifen und umgekehrt. Die Daten werden auf dem Server für dieses Postfach gespeichert und es kann in nachfolgenden Outlook-Sitzungen in allen Formfaktoren auf diese zugegriffen werden, die das Outlook-Add-In unterstützt. 

## Benutzerdefinierte Daten pro Postfach: Roamingeinstellungen


Mithilfe des [RoamingSettings](../reference/outlook/RoamingSettings.md%28Office.15%29.md)-Objekts können Sie Daten angeben, die spezifisch für das Postfach eines Exchange-Benutzers sind. Beispiele solcher Daten sind die persönlichen Daten und Einstellungen des Benutzers. Ihr Mail-Add-In kann auf Roamingeinstellungen zugreifen, wenn das Roaming auf dem entsprechenden Gerät (Desktop, Tablet oder Smartphone) aktiviert ist. 

 Änderungen an diesen Daten werden in einer speicherinterne Kopie der Einstellungen für die aktuelle Outlook-Sitzung gespeichert. Sie müssen alle Roamingeinstellungen nach dem Aktualisieren explizit speichern, damit sie zur Verfügung stehen, sobald der Benutzer Ihr Outlook-Add-In das nächste Mal auf demselben oder einem anderen unterstützten Gerät öffnet.


### Format der Roamingeinstellungen


Die Daten in einem  **RoamingSettings**-Objekt werden als serialisierte JSON-Zeichenfolge (JSON) gespeichert. Im Folgenden sehen Sie ein Beispiel für die Struktur, wobei davon ausgegangen wird, dass die drei Roamingeinstellungen  `add-in_setting_name_0`,  `add-in_setting_name_1` und `add-in_setting_name_2` definiert sind.


```
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### Laden der Roamingeinstellungen


Von einem Mail-Add-In werden Roamingeinstellungen in der Regel im [Office.initialize](http://msdn.microsoft.com/de-de/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx)-Ereignishandler geladen. Im folgenden JavaScript-Codebeispiel wird gezeigt, wie Sie vorhandene Roamingeinstellungen laden und die Werte der beiden Einstellungen „customerName" und „customerBalance" abrufen:


```
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### Erstellen oder Zuweisen einer Roamingeinstellung


Als Fortsetzung des vorhergehenden Beispiels wird nun mit der JavaScript-Funktion  `setAddInSetting` veranschaulicht, wie Sie mithilfe der [RoamingSettings.set](../reference/outlook/RoamingSettings.html%28Office.15%29.md)>-Methode eine Einstellung mit dem Namen  `cookie`und dem heutigen Datum festlegen und die Daten mithilfe der [RoamingSettings.saveAsync](https://dev.outlook.com/reference/add-ins/RoamingSettings.html%28Office.15%29.md)-Methode dauerhaft speichern, indem alle Roamingeinstellungen wieder auf dem Server gespeichert werden. Mit der  **set**-Methode wird die Einstellung soweit noch nicht vorhanden erstellt, und der angegebene Wert wird zugewiesen. Mit der  **saveAsync**-Methode werden Roamingeinstellungen asynchron gespeichert. In diesem Codebeispiel wird die Rückrufmethode  `saveMyAddInSettingsCallback` an **saveAsync** übergeben. Nach Abschluss des asynchronen Aufrufs wird `saveMyAddInSettingsCallback` mithilfe des _asyncResult_-Parameters aufgerufen. Bei diesem Parameter handelt es sich um ein [AsyncResult](https://dev.outlook.com/reference/add-ins/simple-types.md%28Office.15%29.md)-Objekt, welches das Ergebnis und alle Details des asynchronen Aufrufs enthält. Mithilfe des optionalen  _userContext_-Parameters können Sie Statusinformationen vom asynchronen Aufruf an die Rückruffunktion übergeben. 


```
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### Entfernen einer Roamingeinstellung


Als Fortsetzung der vorherigen Beispiele wird mit der JavaScript-Funktion  `removeAddInSetting` veranschaulicht, wie Sie mit der [RoamingSettings.remove](../reference/outlook/RoamingSettings.md%28Office.15%29.md)-Methode die  `cookie`-Einstellung entfernen und alle Roamingeinstellungen wieder auf dem Exchange-Server speichern.


```
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## Benutzerdefinierte Daten pro Element in einem Postfach: benutzerdefinierte Eigenschaften


Sie können Daten für ein bestimmtes Element im Benutzerpostfach mit dem [CustomProperties](../reference/outlook/CustomProperties.md%28Office.15%29.md)-Objekt angeben. Ihr Add-In könnte beispielsweise bestimmte Nachrichten kategorisieren und die Kategorie mithilfe der benutzerdefinierten Eigenschaft  `messageCategory` kennzeichnen. Wenn Ihr Mail-Add-In Termine aus Besprechungsvorschlägen in einer Nachricht erstellt, können Sie mit einer benutzerdefinierten Eigenschaft diese Termine nachverfolgen. Auf diese Weise wird sichergestellt, dass beim zweiten Öffnen der Nachricht durch den Benutzer Ihr Mail-Add-In nicht zum zweiten Mal vorschlägt, den Termin zu erstellen.

Ähnlich wie bei Roamingeinstellungen werden die Änderungen an benutzerdefinierten Eigenschaften in speicherinternen Kopien er Eigenschaften der aktuellen Outlook-Sitzung gespeichert. Um sicherzustellen, dass diese benutzerdefinierten Eigenschaften in der nächsten Sitzung zur Verfügung stehen, speichern Sie alle benutzerdefinierten Eigenschaften auf dem Server.

Auf diese für Add-Ins oder Elemente spezifischen benutzerdefinierten Eigenschaften kann nur mithilfe des  **CustomProperties**-Objekts zugegriffen werden. Diese Eigenschaften weichen von den benutzerdefinierten, MAPI-basierten [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx) im Outlook-Objektmodell und erweiterten Eigenschaften in Exchange-Webdienste ab. Mit dem Outlook-Objektmodell oder EWS kann auf **CustomProperties** nicht zugegriffen werden.

Ein Mail-Add-In kann jedoch mit dem [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)-EWS-Vorgang die MAPI-basierten erweiterten Eigenschaften abrufen. Greifen Sie auf  **GetItem** auf der Serverseite unter Verwendung eines Rückruftokens oder auf der Clientseite unter Verwendung der [mailbox.makeEwsRequestAsync](../reference/outlook/Office.context.mailbox.md%28Office.15%29.md)-Methode zu. Geben Sie in der  **GetItem**-Anforderung die benötigten benutzerdefinierten erweiterten Eigenschaften in einem Eigenschaftensatz an. Ein Mail-Add-In kann auch  **makeEwsRequestAsync** und [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)- und [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)-EWS-Vorgänge zum Erstellen und Ändern erweiterter Eigenschaften verwenden.

Weitere Informationen zu MAPI-Eigenschaften und erweiterten Eigenschaften in Outlook  **UserProperties**und EWS unter Verwendung des Rückruftokens und  **makeEwsRequestAsync** finden Sie im Abschnitt [Zusätzliche Ressourcen](30217d63-7615-4f3f-8618-c91e4e60cd43.md#MailAppCustomProperties_AdditionalResources).


### Verwenden benutzerdefinierter Eigenschaften


Bevor Sie benutzerdefinierte Eigenschaften verwenden, müssen Sie sie durch Aufrufen der [loadCustomPropertiesAsync](../reference/outlook/Office.context.mailbox.item.html%28Office.15%29.md)-Methode laden. Sind bereits benutzerdefinierte Eigenschaften für das aktuelle Element festgelegt, werden sie zu diesem Zeitpunkt vom Exchange-Server geladen. Nach der Erstellung des Eigenschaftenbehälters können Sie mithilfe der [set](https://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)- und [get](https://dev.outlook.com/reference/add-ins/CustomProperties.html%28Office.15%29.md)-Methode benutzerdefinierte Eigenschaften hinzufügen und abrufen. Zum Speichern von Änderungen an Eigenschaftenbehältern müssen Sie die [saveAsync](https://dev.outlook.com/reference/add-ins/CustomProperties.md%28Office.15%29.md)-Methode verwenden, um die Änderungen auf dem Exchange-Server zu speichern.


(../../images  Da benutzerdefinierte Eigenschaften von Outlook für Mac nicht zwischengespeichert werden, können Mail-Add-Ins in Outlook für Mac nicht auf ihre benutzerdefinierten Eigenschaften zugreifen, wenn das Netzwerk des Benutzers nicht verfügbar ist.


### Beispiel für benutzerdefinierte Eigenschaften


Im folgenden Beispiel wird ein einfacher Methodensatz für ein Outlook-Add-In veranschaulicht, in dem benutzerdefinierte Eigenschaften verwendet werden. Sie können dieses Beispiel als Ausgangspunkt für ein eigenes Mail-Add-In nutzen, in dem benutzerdefinierte Eigenschaften verwendet werden. 

Dieses Beispiel umfasst die folgenden Methoden:


- [Office.initialize](http://msdn.microsoft.com/de-de/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx). Initialisiert das Add-In und lädt den benutzerdefinierten Eigenschaftenbehälter vom Exchange-Server.
    
-  **customPropsCallback**. Ruft den vom Server zurückgegebenen benutzerdefinierten Eigenschaftenbehälter ab und speichert ihn zur späteren Verwendung. 
    
-  **updateProperty**: Legt eine bestimmte Eigenschaft fest oder aktualisiert sie und speichert die Änderung dann auf dem Server.
    
-  **removeProperty**. Entfernt eine bestimmte Eigenschaft aus dem Eigenschaftenbehälter und speichert die Entfernung dann auf dem Server.
    



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## Weitere Ressourcen



- [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md)
    
- [Übersicht über Architektur und Features von Outlook-Add-Ins](2cd5641b-492b-4431-8388-7fc589163e9c.md)
    
- [Übersicht über die MAPI-Eigenschaft](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- [Übersicht über Outlook-Eigenschaften](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [Aufrufen von Webdiensten aus einem Outlook-Add-In](0049645e-9af1-46c7-a22e-fe6c650b134e.md)
    
- [Eigenschaften und erweiterten Eigenschaften in EWS in Exchange](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- [Eigenschaftensätze und Antwort shapes in EWS in Exchange](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


