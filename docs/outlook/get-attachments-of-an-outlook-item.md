
# Abrufen von Anlagen eines Outlook-Elements vom Server
Verwenden eines Outlook-Add-Ins zum Aktivieren eines Dienstes, der Anlagen von einem Exchange-Server abruft. 

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Von einem Outlook-Add-In können die Anlagen eines Elements nicht direkt an den auf Ihrem Server ausgeführten Remotedienst übergeben werden. Stattdessen kann von dem Outlook-Add-In die Anlagen-API verwendet werden, um Informationen zu den Anlagen an den Remotedienst zu senden. Der Dienst kann daraufhin direkt den Exchange-Server kontaktieren, um die Anlagen abzurufen.

Zum Senden von Anlageninformationen an den Remotedienst werden folgende Eigenschaften und die folgende Funktion verwendet:


- [Office.context.mailbox.ewsUrl](../reference/outlook/Office.context.mailbox.md%28Office.15%29.md)-Eigenschaft – Stellt die URL von Exchange-Webdienste (EWS) auf dem Exchange Server bereit, auf dem das Postfach gehostet wird. Ihr Dienst verwendet diese URL, um die [verwaltete EWS-API](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)-Methode [ExchangeService.GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) oder den EWS-Vorgang [GetAttachment](http://msdn.microsoft.com/en-us/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) aufzurufen.
    
- Eigenschaft [Office.context.mailbox.item.attachments](../reference/outlook/Office.context.mailbox.item.html%28Office.15%29.md) – Ruft ein Array von [AttachmentDetails](https://dev.outlook.com/reference/add-ins/simple-types.md%28Office.15%29.md)-Objekten ab (jeweils ein Objekt für jede Anlage des Elements).
    
- Funktion [Office.context.mailbox.getCallbackTokenAsync](../reference/outlook/Office.context.mailbox.md%28Office.15%29.md) – Nimmt einen asynchronen Anruf beim Exchange-Server vor, auf dem das Postfach gehostet ist, um ein Rückruftoken zu erhalten, das der Server dann zurück an den Exchange-Server sendet, um eine Anforderung für eine Anlage zu authentifizieren.
    

## Verwenden der Anlagen-API


Wenn Sie die Anlagen-API zum Abrufen von Anlagen aus einem Exchange-Postfach verwenden möchten, gehen Sie wie folgt vor: 


1. Zeigen Sie das Outlook-Add-In an, wenn eine Nachricht oder ein Termin eine Anlage enthält.
    
2. Rufen Sie das Rückruftoken vom Exchange-Server ab.
    
3. Senden Sie Rückruftoken und Anlageninformationen an den Remotedienst.
    
4. Rufen Sie die Anlagen mithilfe der Methode  **ExchangeService.GetAttachments** oder dem Vorgang **GetAttachment** vom Exchange-Server ab.
    
Die einzelnen Schritte werden in den folgenden Abschnitten ausführlich und unter Verwendung von Code aus dem Beispiel [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) erläutert.


(../../images  Der Code in diesen Beispielen wurde gekürzt, um die Anlageninformationen hervorzuheben. Das Beispiel enthält zusätzlichen Code zum Authentifizieren des Outlook-Add-Ins beim Remoteserver sowie zum Verwalten des Anforderungszustands.


### Aktivieren des Outlook-Add-Ins


Sie können eine [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)-Regel in der Manifestdatei des Outlook-Add-Ins verwenden, um Ihr Outlook-Add-In anzuzeigen, wenn das Element Anlagen enthält, siehe folgendes Beispiel.


```XML
<Rule xsi:type="ItemHasAttachment" />
```


### Abrufen eines Rückruftokens


Das [Office.context.mailbox](../reference/outlook/Office.context.mailbox.md%28Office.15%29.md)-Objekt stellt die  **getCallbackTokenAsync**-Funktion zum Abrufen eines Tokens bereit, mit dessen Hilfe sich der Remoteserver beim Exchange Server authentifizieren kann. Der folgende Code zeigt eine Funktion in einem Add-In, die die asynchrone Anforderung zum Erhalten des Rückruftokens startet, sowie die Rückruffunktion, von der die Antwort abgerufen wird. Das Rückruftoken wird im Dienstanforderungsobjekt gespeichert, das im nächsten Abschnitt definiert wird.


```
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "" {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
};
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "success") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
};
```


### Senden von Anlageninformationen an den Remotedienst


Die genaue Vorgehensweise zum Senden von Informationen an den Remotedienst ist abhängig davon, welcher Dienst von Ihrem Outlook-Add-In aufgerufen wird. In diesem Beispiel handelt es sich bei dem Remotedienst um eine mithilfe von Visual Studio 2013 erstellte Web-API-Anwendung. Die Anlageninformationen werden vom Remotedienst in einem JSON-Objekt erwartet. Der folgende Code dient zum Initialisieren eines Objekts mit den Anlageninformationen:


```
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
serviceRequest = new Object();
serviceRequest.attachmentToken = "";
serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
serviceRequest.attachments = new Array();
```

Die Eigenschaft  `Office.context.mailbox.item.attachments` enthält eine Sammlung von Objekten des Typs **AttachmentDetails** (jeweils ein Objekt für jede Anlage des Elements). In den meisten Fällen kann von dem Outlook-Add-In lediglich die Anlagen-ID-Eigenschaft eines Objekts vom Typ **AttachmentDetails** an den Remotedienst übergeben werden. Sollte der Remotedienst weitere Informationen zur Anlage benötigen, können Sie das Objekt **AttachmentDetails** (ganz oder teilweise) übergeben. Der folgende Code dient zum Definieren einer Methode, die das gesamte Array vom Typ **AttachmentDetails** in das Objekt `serviceRequest` einfügt und eine Anforderung an den Remotedienst sendet:




```
    function makeServiceRequest() {
      // Format the attachment details for sending.
      for (var i = 0; i < mailbox.item.attachments.length; i++) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i].$0_0));
      }

      $.ajax({
        url: '../../api/Default',
        type: 'POST',
        data: JSON.stringify(serviceRequest),
        contentType: 'application/json;CHARSET=Windows-1252'
      }).done(function (response) {
        if (!response.isError) {
          var names = "<h2>Attachments processed using " +
                        serviceRequest.service +
                        ": " +
                        response.attachmentsProcessed +
                        "</h2>";
          for (i = 0; i < response.attachmentNames.length; i++) {
            names += response.attachmentNames[i] + "<br />";
          }
          document.getElementById("names").innerHTML = names;
        } else {
          app.showNotification("Runtime error", response.message);
        }
      }).fail(function (status) {

      }).always(function () {
        $('.disable-while-sending').prop('disabled', false);
      })
    };

```


### Abrufen der Anlagen vom Exchange-Server


Ihr Remotedienst kann entweder die verwaltete EWS API-Methode [GetAttachments](http://msdn.microsoft.com/de-de/library/office/dn600509%28v=exchg.80%29.aspx) oder den EWS-Vorgang [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) verwenden, um Anlagen vom Server abzurufen. Die Dienstanwendung benötigt zwei Objekte um die JSON-Zeichenfolge in .NET Framework-Objekte zu deserialisieren, die auf dem Server verwendet werden können. Der folgende Code zeigt die Definitionen der Deserialisierungsobjekte:


```C#



namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```


#### Verwenden der verwalteten EWS-API zum Abrufen der Anlagen

Wenn Sie die [verwaltete EWS-API](http://go.microsoft.com/fwlink/?LinkID=255472) im Remotedienst einsetzen, können Sie die [GetAttachments](http://msdn.microsoft.com/de-de/library/office/dn600509%28v=exchg.80%29.aspx)-Methode verwenden, die eine EWS-SOAP-Anforderung erstellt, sendet und empfängt, um die Anlagen abzurufen. Microsoft empfiehlt die verwaltete EWS-API, da sie weniger Codezeilen erfordert und eine intuitivere Benutzeroberfläche für EWS-Anrufe bereitstellt. Der folgende Code sendet eine Anforderung, um alle Anlagen abzurufen, und gibt die Anzahl und die Namen der verarbeiteten Anlagen zurück.


```C#
    private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      // Create an ExchangeService object, set the credentials and the EWS URL.
      ExchangeService service = new ExchangeService();
      service.Credentials = new OAuthCredentials(request.attachmentToken);
      service.Url = new Uri(request.ewsUrl);

      var attachmentIds = new List<string>();

      foreach (AttachmentDetails attachment in request.attachments)
      {
        attachmentIds.Add(attachment.id);
      }

      // Call the GetAttachments method to retrieve the attachments on the message.
      // This method results in a GetAttachments EWS SOAP request and response
      // from the Exchange server.
      var getAttachmentsResponse =
        service.GetAttachments(attachmentIds.ToArray(),
                               null,
                               new PropertySet(BasePropertySet.FirstClassProperties,
                                               ItemSchema.MimeContent));

      if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
      {
        foreach (var attachmentResponse in getAttachmentsResponse)
        {
          attachmentNames.Add(attachmentResponse.Attachment.Name);

          // Write the content of each attachment to a stream.
          if (attachmentResponse.Attachment is FileAttachment)
          {
            FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
            Stream s = new MemoryStream(fileAttachment.Content);
            // Process the contents of the attachment here.
          }

          if (attachmentResponse.Attachment is ItemAttachment)
          {
            ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
            Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
            // Process the contents of the attachment here.
          }

          attachmentsProcessedCount++;
        }
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }


```


#### Verwenden von EWS zum Abrufen von Anlagen

Wenn Sie EWS im Remotedienst verwenden, müssen Sie eine SOAP-Anforderung vom Typ [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) erstellen, um die Anlagen vom Exchange-Server abzurufen. Der folgende Code gibt eine Zeichenfolge zurück, die die SOAP-Anforderung bereitstellt. Der Remotedienst verwendet die **String.Format** -Methode, um die Anlagen-ID einer Anlage in die Zeichenfolge einzufügen.


```C#
    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";

```

Die folgende Methode verwendet die EWS-Anforderung  **GetAttachment**, um die Anlagen vom Exchange-Server abzurufen. Diese Implementierung sendet eine einzige Anforderung pro Anlage und gibt die Anzahl der verarbeiteten Anlagen zurück. jede Antwort wird in einer separaten  **ProcessXmlResponse**-Methode verarbeitet, die im Folgenden definiert wird.




```C#
    private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      foreach (var attachment in request.attachments)
      {
        // Prepare a web request object.
        HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
        webRequest.Headers.Add("Authorization",
          string.Format("Bearer {0}", request.attachmentToken));
        webRequest.PreAuthenticate = true;
        webRequest.AllowAutoRedirect = false;
        webRequest.Method = "POST";
        webRequest.ContentType = "text/xml; CHARSET=Windows-1252";

        // Construct the SOAP message for the GetAttachment operation.
        byte[] bodyBytes = Encoding.UTF8.GetBytes(
          string.Format(GetAttachmentSoapRequest, attachment.id));
        webRequest.ContentLength = bodyBytes.Length;

        Stream requestStream = webRequest.GetRequestStream();
        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        requestStream.Close();

        // Make the request to the Exchange server and get the response.
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        // If the response is okay, create an XML document from the reponse
        // and process the request.
        if (webResponse.StatusCode == HttpStatusCode.OK)
        {
          var responseStream = webResponse.GetResponseStream();

          var responseEnvelope = XElement.Load(responseStream);

          // After creating a memory stream containing the contents of the 
          // attachment, this method writes the XML document to the trace output.
          // Your service would perform it's processing here.
          if (responseEnvelope != null)
          {
            var processResult = ProcessXmlResponse(responseEnvelope);
            attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

          }

          // Close the response stream.
          responseStream.Close();
          webResponse.Close();

        }
        // If the response is not OK, return an error message for the 
        // attachment.
        else
        {
          var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
            "Error message: {1}.", attachment.name, webResponse.StatusDescription);
          attachmentNames.Add(errorString);
        }
        attachmentsProcessedCount++;
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }

```

Jede Antwort vom Vorgang  **GetAttachment** wird an die **ProcessXmlResponse**-Methode gesendet. Die Methode prüft dann die Antwort auf Fehler. Findet sie keinen Fehler, werden die Dateianlagen und die Objektanlagen verarbeitet. Die  **ProcessXmlResponse**-Methode verarbeitet die Anlage.




```C#
    // This method processes the response from the Exchange server.
    // In your application the bulk of the processing occurs here.
    private string ProcessXmlResponse(XElement responseEnvelope)
    {
      // First, check the response for web service errors.
      var errorCodes = from errorCode in responseEnvelope.Descendants
                       ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                       select errorCode;
      // Return the first error code found.
      foreach (var errorCode in errorCodes)
      {
        if (errorCode.Value != "NoError")
        {
          return string.Format("Could not process result. Error: {0}", errorCode.Value);
        }
      }

      // No errors found, proceed with processing the content.
      // First, get and process file attachments.
      var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                            select fileAttachment;
      foreach(var fileAttachment in fileAttachments)
      {
        var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
        var fileData = System.Convert.FromBase64String(fileContent.Value);
        var s = new MemoryStream(fileData);
        // Process the file attachment here. 
      }

      // Second, get and process item attachments.
      var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                            ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                            select itemAttachment;
      foreach(var itemAttachment in itemAttachments)
      {
        var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
        if (message != null)
        {
         // Process a message here.
          break;
        }
        var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
        if (calendarItem != null)
        {
          // Process calendar item here.
          break;
        }
        var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
        if (contact != null)
        {
          // Process contact here.
          break;
        }
        var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
        if (task != null)
        {
          // Process task here.
          break;
        }
        var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
        if (meetingMessage != null)
        {
          // Process meeting message here.
          break;
        }
        var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
        if (meetingRequest != null)
        {
          // Process meeting request here.
          break;
        }
        var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
        if (meetingResponse != null)
        {
          // Process meeting response here.
          break;
        }
        var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
        if (meetingCancellation != null)
        {
          // Process meeting cancellation here.
          break;
        }
      }
     
      return string.Empty;
    }

```


## Zusätzliche Ressourcen



- [Erstellen von Outlook-Add-Ins für Leseformulare](c8680c4e-3068-4e13-8537-2a96c614193e.md)
    
- [Erkunden von verwalteter EWS-API, EWS und Webdiensten in Exchange](http://msdn.microsoft.com/library/0bc6f81d-cc10-42b0-ba5d-6f22ff55d51c%28Office.15%29.aspx)
    
- [Erste Schritte mit verwalteten EWS-API-Clientanwendungen](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)
    
- [Outlook-Power-Hour_Codebeispiele](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples):  `MyAttachments` und `AttachmentsDemo`
    
