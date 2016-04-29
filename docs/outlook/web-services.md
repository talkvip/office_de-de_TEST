
# Aufrufen von Webdiensten aus einem Outlook-Add-In
Erfahren Sie mehr darüber, wie Sie die Exchange Webdienste (EWS) oder andere Webdienste von einem Outlook-Add-In für Outlook aufrufen können.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Ihr Mail-Add-in kann Exchange-Webdienste (EWS) auf einem Computer verwenden, auf dem Exchange Server 2013 ausgeführt wird, einen Webdienst, der auf dem Server, über den der Quellspeicherort für die Benutzeroberfläche des Add-Ins bereitgestellt wird, oder im Internet verfügbar ist. Dieser Artikel enthält ein Beispiel, das zeigt, die ein Outlook-Add-In Informationen von EWS abrufen kann.

Die zum Aufrufen eines Webdiensts gewählte Methode basiert auf dem Speicherort des Webdiensts. In der folgenden Tabelle sind die Methoden für unterschiedliche Speicherorte von Webdiensten angegeben.


**Tabelle 1. Möglichkeiten zum Aufruf von Webdiensten aus einem Outlook-Add-In**


|**Speicherort des Webdiensts**|**Möglichkeit zum Aufruf des Webdienstes**|
|:-----|:-----|
|Der Exchange Server dient als Host des Clientpostfachs.|Verwenden Sie die [mailbox.makeEwsRequestAsync](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode, um EWS-Vorgänge aufzurufen, die Add-Ins unterstützen. Der Exchange Server hostet das Postfach hostet und stellt auch die Exchange-Webdienste zur Verfügung.|
|Der Webserver stellt den Quellspeicherort der Add-in-Benutzeroberfläche bereit.|Rufen Sie den Webdienst mithilfe standardmäßiger JavaScript-Techniken auf. Der JavaScript im UI-Frame wird im Kontext des Webservers ausgeführt, der die Benutzeroberfläche (UI) bereitstellt. Deshalb können Webdienste auf diesem Server aufgerufen werden, ohne einen Fehler aufgrund von websiteübergreifendem Skripting zu verursachen.|
|Alle anderen Speicherorte.|Erstellen Sie einen Proxy für den Webdienst auf dem Webserver, der den Quellspeicherort der Benutzeroberfläche bereitstellt. Wenn Sie keinen Proxy einrichten, kann Ihr Add-In aufgrund von Fehlern wegen websiteübergreifendem Skripting nicht ausgeführt werden. Eine Möglichkeit zum Bereitstellen eines solchen Proxys bietet JSON/P. Weitere Informationen finden Sie unter [Datenschutz und Sicherheit bei Office-Add-Ins](87c59a88-10e2-4c88-b6a8-736bd356e5f8.md).|

## Die makeEwsRequestAsync-Methode zum Zugriff auf EWS-Vorgänge verwenden
<a name="mod_off15_appscope_CallingWebServices_UsingEWS"> </a>

Sie können die [mailbox.makeEwsRequestAsync](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md)-Methode verwenden, um eine EWS-Anforderung an den Exchange-Server zu stellen, auf dem das Postfach des Benutzers gehostet wird.

Exchange-Webdienste unterstützen unterschiedliche Vorgänge auf einem Exchange Server, z. B. Vorgänge auf Elementebene zum Kopieren, Suchen, Aktualisieren oder Senden eines Elements und Vorgänge auf Ordnerebene zum Erstellen, Abrufen oder Aktualisieren eines Ordners. Erstellen Sie zum Durchführen eines EWS-Vorgangs eine XML-SOAP-Anforderung für den betreffenden Vorgang. Nach Abschluss des Vorgangs erhalten Sie eine XML-SOAP-Antwort mit den Daten, die für den Vorgang relevant sind. EWS-SOAP-Anforderungen und -Antworten basieren auf dem Schema, das in der Datei "Messages.xsd" definiert ist. Wie andere EWS-Schemadateien auch, ist die Datei "Message.xsd" in dem virtuellen IIS-Verzeichnis enthalten, von dem der Exchange Server gehostet wird. 

Stellen Sie bitte Folgendes bereit, um die  **makeEwsRequestAsync**-Methode zur Initiierung eines EWS-Vorgang zu verwenden:


- Die XML für die SOAP-Anforderung für den EWS-Vorgang, als Argument für den  _data_ Parameter
    
- Eine Rückrufmethode (als  _callback_-Argument)
    
- Sämtliche optionale Eingabedaten für die Rückrufmethode (als  _userContext_-Argument)
    
Nach Abschluss der EWS-SOAP-Anforderung wird von Outlook die Rückrufmethode mit einem Argument aufgerufen, bei dem es sich um ein [AsyncResult](https://dev.outlook.com/reference/add-ins/simple-types.html%28Office.15%29.md)-Objekt handelt. Die Rückrufmethode kann auf zwei Eigenschaften des  **AsyncResult**-Objekts zugreifen: auf die  **value**-Eigenschaft , in der die XML-SOAP-Antwort des EWS-Vorgangs enthalten ist, und optional auf die  **asyncContext**-Eigenschaft , in der alle als  **userContext**-Parameter übergebenen Daten enthalten sind. Normalerweise analysiert die Rückrufmethode dann die XML-Daten in der SOAP-Anforderung, um relevante Informationen abzurufen, und verarbeitet diese Informationen entsprechend.


## Tipps für die Analyse von EWS-Antworten
<a name="mod_off15_appscope_CallingWebServices_Tips"> </a>

Bei der Analyse einer SOAP-Antwort von einem EWS-Vorgang sind die folgenden Browser-abhängigen Probleme zu beachten:


- Geben Sie ein Präfix für einen Tagnamen an, wenn Sie die DOM-Methode  **getElementsByTagName** verwenden, damit der Internet Explorer unterstützt wird.
    
    Das  **getElementsByTagName** verhält sich anders, je nach Browsertyp. Eine EWS-Antwort kann beispielsweise die folgenden XML-Daten beinhalten (formatiert und abgekürzt zu Ansichtszwecken):
    


  ```XML
  <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
PropertyName="MyProperty" 
PropertyType="String"/>
<t:Value>{
...
}</t:Value></t:ExtendedProperty>
  ```


    Mit dem folgenden Code können in einem Browser wie Chrome die XML-Daten, die von  **ExtendedProperty**-Tags eingeschlossen sind, abgerufen werden:
    


  ```
  var mailbox = Office.context.mailbox;
mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
    var response = $.parseXML(result.value);
    var extendedProps = response.getElementsByTagName("ExtendedProperty");
  ```


    Im Internet Explorer muss das  `t:`-Präfix des Tagnamens auf folgende Weise eingeschlossen sein:
    


  ```
  var mailbox = Office.context.mailbox;
mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
    var response = $.parseXML(result.value);
    var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
  ```

- Verwenden Sie die DOM-Eigenschaft  **textContent**, um Inhalte eines Tags in einer EWS-Antwort, wie im Folgenden dargestellt, abzurufen:
    
  ```
  content = $.parseJSON(value.textContent);
  ```


    Weitere Eigenschaften wie  **innerHTML** funktionieren möglicherweise nicht im Internet Explorer für einige Tags in einer EWS-Antwort.
    

## Beispiel
<a name="mod_off15_appscope_CallingWebServices_Example"> </a>

Im folgenden Beispiel wird  **makeEwsRequestAsync** aufgerufen, um den[GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)-Vorgang zum Abrufen des Betreffs eines Elements zu verwenden. Dieses Beispiel enthält die folgenden drei Funktionen: 


-  `getSubjectRequest` – Verwendet eine Element-ID als Eingabe und gibt die XML-Daten für die SOAP-Anforderung zum Aufrufen von **GetItem** für das angegebene Element zurück.
    
-  `sendRequest`— Ruft  `getSubjectRequest`auf, um die SOAP-Anforderung für das ausgewählte Element abzurufen. Anschließend werden die SOAP-Anforderung und die  `callback`-Methode an  **makeEwsRequestAsync** übergeben, um den Betreff des angegebenen Elements abzurufen.
    
-  `callback` – Verarbeitet die SOAP-Antwort, in der alle Informationen zum Betreff sowie die anderen Informationen zum angegebenen Element enthalten sind.
    

```
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## Von Add-Ins unterstützte EWS-Vorgänge
<a name="mod_off15_appscope_CallingWebServices_SupportedEWS"> </a>

Outlook-Add-Ins können auf eine Teilmenge der Vorgänge, die in EWS über die  **makeEwsRequestAsync**-Methode verfügbar sind, zugreifen. Falls Sie nicht mit EWS-Vorgängen vertraut sind und nicht wissen, wie Sie die  **makeEwsRequestAsync**-Methode verwenden können, um auf einen Vorgang zuzugreifen, beginnen Sie mit einem SOAP-Anfordungsbeispiel, um Ihr  _data_-Argument anzupassen. Im Folgenden wird beschrieben, wie Sie die  **makeEwsRequestAsync** -Methode verwenden können:


1. Ersetzen Sie in den XML-Daten alle Element-IDs und relevanten EWS-Vorgangsattribute durch die jeweils passenden Werte.
    
2. Ordnen Sie die SOAP-Anforderung als Argument für den  _data_-Parameter von  **makeEwsRequestAsync** an.
    
3. Geben Sie eine Rückrufmethode an, und rufen Sie  **makeEwsRequestAsync** auf.
    
4. Überprüfen Sie in der Rückrufmethode die Ergebnisse des Vorgangs in der SOAP-Antwort. Nutzen Sie diese Ergebnisse je nach Bedarf.
    
5. Verwenden Sie die Ergebnisse des EWS-Vorgangs je nach Bedarf.
    
Die folgende Tabelle führt die EWS-Vorgänge auf, die von Add-Ins unterstützt werden. Klicken Sie für jeden Vorgang auf den entsprechenden Link, um Beispiele von SOAP-Anforderungen und -Antworten anzuzeigen. Weitere Informationen über EWS-Vorgänge finden Sie unter [EWS-Operationen in Exchange](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx).


**Tabelle 2. Unterstützte EWS-Vorgänge**


|**EWS-Vorgänge**|**Beschreibung**|
|:-----|:-----|
|[CopyItem Operation](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|Kopiert die angegebenen Elemente und fügt die neuen Elemente in einen bestimmten Ordner im Exchange-Speicher ein.|
|[CreateFolder Operation](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|Erstellt Ordner am angegebenen Speicherort im Exchange-Speicher.|
|[CreateItem Operation](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|Erstellt die angegebenen Elemente im Exchange-Speicher.|
|[FindConversation Operation](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|Führt eine Liste mit Unterhaltungen im angegebenen Ordner im Exchange-Speicher auf.|
|[FindFolder Operation](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|Sucht nach Unterordnern eines angegebenen Ordners und gibt Eigenschaften zurück, mit denen die Unterordner beschrieben werden.|
|[FindItem Operation](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|Identifiziert Elemente, die in einem angegebenen Ordner im Exchange-Speicher enthalten sind.|
|[GetConversationItems operation](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|Ruft einen weiteren Elementesatz auf, der in einer Unterhaltung in Knoten angeordnet ist.|
|[GetFolder Operation](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|Ruft die angegebenen Eigenschaften und den Inhalt von Ordnern aus dem Exchange-Speicher ab.|
|[GetItem Operation](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|Ruft die angegebenen Eigenschaften und den Inhalt von Elementen aus dem Exchange-Speicher ab.|
|[MarkAsJunk Operation](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|Verschiebt E-Mail-Nachrichten in den Ordner  **Junk-E-Mail** und fügt Absender der Nachrichten entsprechend der Liste der blockierten Absender hinzu bzw. entfernt diese daraus.|
|[MoveItem Operation](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|Verschiebt Elemente in einen gemeinsamen Zielordner im Exchange-Speicher.|
|[SendItem Operation](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|Sendet E-Mail-Nachrichten, die sich im Exchange-Speicher befinden.|
|[UpdateFolder-Vorgang](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|Ändert die Eigenschaften bestehender Ordner im Exchange-Speicher.|
|[UpdateItem Operation](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|Ändert die Eigenschaften bestehender Elemente im Exchange-Speicher.|

## Authentifizierungs- und Berechtigungsabwägungen für die makeEwsRequestAsync-Methode
<a name="mod_off15_appscope_CallingWebServices_AuthAndPerms"> </a>

Bei Verwendung der  **makeEwsRequestAsync**-Methode wird die Anforderung mithilfe der Anmeldeinformationen aus dem E-Mail-Konto des aktuellen Benutzers authentifiziert. Die  **makeEwsRequestAsync**-Methode verwaltet die Anmeldeinformationen für Sie, damit Sie in Ihrer Anforderung keine Anmeldeinformationen zur Authentifizierung angeben müssen.


 **Hinweis**  Der Serveradministrator muss das [New-WebServicesVirtualDirctory](http://technet.microsoft.com/de-de/library/bb125176.aspx) oder das[Set-WebServicesVirtualDirecory](http://technet.microsoft.com/de-de/library/aa997233.aspx)-Cmldet verwenden, um den  _OAuthAuthentication_-Parameter im EWS-Verzeichnis des Clientzugriffsservers auf  **true** festzulegen, damit die **makeEwsRequestAsync**-Methode zum Erstellen von EWS-Anforderungen aktiviert wird.

Für das Add-In muss die  **ReadWriteMailbox**-Berechtigung im dazugehörigen Add-In-Manifest angegeben sein, damit die Methode  **makeEwsRequestAsync** verwendet werden kann. Informationen zur Verwendung der **ReadWriteMailbox**-Berechtigung finden Sie im Abschnitt [ReadWriteMailbox-Berechtigung](../outlook/understanding-outlook-add-in-permissions.md#olowa15conagave_permmodelreadwrite) unter[Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../outlook/understanding-outlook-add-in-permissions.md).


## Zusätzliche Ressourcen
<a name="mod_off15_appscope_CallingWebServices_AdditionalRsc"> </a>


- [Outlook-Add-Ins](71e64bc9-e347-4f5d-8948-0a47b5dd93e6.md)
    
- [Datenschutz und Sicherheit bei Office-Add-Ins](87c59a88-10e2-4c88-b6a8-736bd356e5f8.md)
    
- [Behandeln von Richtlinieneinschränkungen aufgrund desselben Ursprungs in Office-Add-ins](36c800ae-1dda-4ea8-a558-37c89ffb161b.md)
    
- [EWS-Referenz für Exchange](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- [Mail-Apps für Outlook und EWS in Exchange](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
Im folgenden Abschnitt finden Sie Informationen zum Erstellen von Back-End-Diensten für Add-Ins mit der ASP.NET Web-API:


- [Erstellen eines Webdiensts für ein Office-Add-in mit der ASP.NET Web-API](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [Grundlagen zum Erstellen eines HTTP-Diensts mit der ASP.NET Web-API](http://www.asp.net/web-api)
    
