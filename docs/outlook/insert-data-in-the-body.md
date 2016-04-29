
# Einfügen von Daten in den Textkörper bei der Erstellung eines Termins oder einer Nachricht in Outlook
Erfahren Sie mehr darüber, wie Daten in den Textkörper eines Termins oder einer Nachricht, die der Benutzer in einem Outlook-Add-In verfasst, eingefügt werden.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Einfügen von Daten in den Textkörper eines Erstellformulars


Aysynchrone Methoden ([Body.getAsync](../reference/outlook/Body.html%28Office.15%29.md), [Body.getTypeAsync](../reference/outlook/Body.html%28Office.15%29.md), [Body.prependAsync](https://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md), [Body.setAsync](https://dev.outlook.com/reference/add-ins/Body.md%28Office.15%29.md) und [Body.setSelectedDataAsync](https://dev.outlook.com/reference/add-ins/Body.md%28Office.15%29.md)) können zum Abrufen des Textkörpertyps und zum Einfügen von Daten in den Textkörper eines Termins oder einer Nachricht verwendet werden, die der Benutzer erstellt. Diese asynchronen Methoden sind nur für Verfassen-Add-Ins verfügbar. Um diese Methoden zu verwenden, müssen Sie sicherstellen, dass Sie das Add-in-Manifest ordnungsgemäß enigerichtet haben, sodass Outlook die Add-Ins in einem Formular zum Verfassen aktiviert, wie im Abschnitt [Einrichten von Outlook-Add-Ins für Erstellformulare](e4126e58-4ddc-4891-9f19-aa6c1a258027.md#mod_off15_CreatingForCompose_SettingUp) von [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md) beschrieben.

In Outlook kann ein Benutzer eine Nachricht im Text-, HTML- oder Rich Text-Format (RTF) erstellen, und kann einen Termin im HTML-Format erstellen. Vor dem Einfügen sollten Sie das unterstützte Elementformat durch Aufrufen von  **getTypeAsync** prüfen, da möglicherweise zusätzliche Schritte notwendig sind. Der Wert, der von **getTypeAsync** zurückgegeben wird, hängt vom ursprünglichen Elementformat ab, sowie von der Unterstützung für das Betriebssystem des Geräts und den Host zum Bearbeiten des HTML-Formats1. Legen Sie dann den  _coercionType_-Parameter von  **prependAsync** oder **setSelectedDataAsync** dementsprechend fest2 , um Daten, wie in der folgenden Tabelle dargestellt, einzufügen. Wenn Sie kein Argument angeben, nehmen **prependAsync** und **setSelectedDataAsync** an, dass die einzufügenden Daten im Textformat vorliegen.



|**Einzufügende Daten**|**Von getTypeAsync zurückgegebenes Elementformat**|**Verwenden Sie diesen coercionType**|
|:-----|:-----|:-----|
|Text|Text1|Text|
|HTML|Text1|Text2|
|Text|HTML|Text/HTML|
|HTML|HTML |HTML|
1Auf Tablets und Smartphones wird von  **getTypeAsync** **Office.MailboxEnums.BodyType.Text** zurückgegeben, wenn das Betriebssystem oder der Host das Bearbeiten von Elementen, die ursprünglich mit HTML, im HTML-Format, erstellt wurden, nicht unterstützt.

2Wenn die einzufügenden Daten im HTML-Format vorhanden sind, und von  **getTypeAsync** ein Texttyp für dieses Element zurückgegeben wird, organisieren Sie die Daten als Text neu und fügen Sie sie mit **Office.MailboxEnums.BodyType.Text** als _coercionType_ ein. Wenn Sie die HTML-Daten mit einem Koersionstyp "Text" einfügen, zeigt der Host die HTML-Tags als Text an. Wenn Sie versuchen, HTML-Daten mit **Office.MailboxEnums.BodyType.Html** als _coercionType_ einzufügen, tritt ein Fehler auf.

Neben  _coercionType_, wie bei den meisten asynchronen Methoden in JavaScript-API für Office, verwenden  **getTypeAsync**,  **prependAsync** und **setSelectedDataAsync** weitere optionale Eingabeparameter. Weitere Informationen zum Angeben dieser optionalen Eingabeparameter finden Sie unter [Übergeben von optionalen Parametern an asynchrone Methoden](http://msdn.microsoft.com/de-de/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters) in [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md).


## So werden Daten an der aktuellen Mauszeigerposition eingefügt


In diesem Abschnitt wird ein Codebeispiel gegeben, das  **getTypeAsync** verwendet, um den Textkörper des Elements, das gerade verfasst wird, zu prüfen, und mit dem anschließend mithilfe von **setSelectedDataAsync** Daten an der aktuellen Mauszeigerposition eingefügt werden.

Sie können eine Rückrufmethode und optionale Eingabeparameter an  **getTypeAsync** übergeben und Status und Ergebnisse im _asyncResult_-Ausgabeparameter abrufen. Wenn die Methode erfolgreich ist, können Sie den Textkörpertyp in der [AsyncResult.value](http://msdn.microsoft.com/en-us/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx)-Eigenschaft abrufen, die entweder "text" oder "html" ist.

Sie müssen eine Datenfolge als Eingabeparameter an  **setSelectedDataAsync** übergeben. Je nach Typ des Elementkörpers können Sie diese Datenfolge im Text- oder HTML-Format angeben. Wie oben bereits erwähnt, können Sie den Typ der Daten, die in den _coercionType_-Parameter eingegeben werden sollen, wahlweise angeben. Zudem können Sie eine Rückrufmethode und jeden ihrer Parameter als optionale Eingabeparameter bereitstellen.

Wenn der Benutzer den Mauszeiger nicht im Textkörper platziert hat, fügt  **setSelectedDataAsync** die Daten zu Beginn des Textkörpers ein. Wenn der Benutzer im Textkörper Text ausgewählt hat, ersetzt **setSelectedDataAsync** den ausgewählten Text mit den von Ihnen angegebenen Daten. Beachten Sie, dass **setSelectedDataAsync** versagen kann, wenn der Benutzer gleichzeitig die Mauszeigerposition ändert und das Element verfasst. Die maximal zulässige Anzahl Zeichen, die Sie mit einem Mal eingeben können, ist auf 1000000 Zeichen beschränkt.

Dieses Codebeispiel setzt eine Regel in dem Add-In voraus, die das Add-In in einem Formular zum Verfassen für einen Termin oder eine Nachricht wie unten dargestellt aktiviert.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## So werden Daten zu Beginn des Textkörpers eines Elements eingefügt


Alternativ können Sie  **prependAsync** verwenden, um zu Beginn des Elementtextkörpers Daten einzufügen und die aktuelle Mauszeigerposition zu irgnorieren. Anders als der Einfügepunkt verhalten sich **prependAsync** und **setSelectedDataAsync** ähnlich:


- Wenn Sie in einem Nachrichtenkörper HTML-Daten voranstellen, sollten Sie zunächst prüfen, ob der Typ des Nachrichtenkörpers verhindert, dass HTML-Daten einer Nachricht im Textformat vorangestellt werden können.
    
- Stellen Sie die folgenden Daten als Eingabedaten für  **prependAsync** bereit: eine Datenfolge, entweder im Textformat oder im HTML-Format und wahlweise das Format der Daten, die eingefügt werden sollen, eine Rückrufmethode und ihre Parameter.
    
- Die maximal zulässige Anzahl Zeichen, die Sie mit einem Mal voranstellen können, ist auf 1000000 Zeichen beschränkt.
    
Der folgende JavaScript-Code ist Teil eines Beispiel-Add-Ins, die in Formularen zum Verfassen von Terminen und Nachrichten aktiviert wird. Im Beispiel wird  **getTypeAsync** abgerufen, um den Typ des Textkörpers des Elements zu prüfen; außerdem werden HTML-Daten zu Beginn des Textkörpers eingegeben, wenn das Element ein Termin oder eine HTML-Nachricht ist; andernfalls werden die Daten im Textformat eingegeben.




```
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Zusätzliche Ressourcen



- [Abrufen und Festlegen von Elementdaten in einem Erstellformular in Outlook](60901b90-2472-4132-9b5e-0749e2e33cbe.md)
    
- [Abrufen und Festlegen von Outlook-Elementdaten in Formularen zum Lesen oder Verfassen](081ef8af-fc0b-4b1c-86e0-4ce8392e9886.md)
    
- [Erstellen von Outlook-Add-Ins für Formulare zum Verfassen](e4126e58-4ddc-4891-9f19-aa6c1a258027.md)
    
- [Asynchrone Programmierung in Office-Add-Ins](7fe6bb42-3178-4d96-85f5-af5caea7b950.md)
    
- [Abrufen, Festlegen oder Hinzufügen von Empfängern beim Erstellen eines Termins oder einer Nachricht in Outlook](c92f62f5-cd93-49c2-9c74-d9b195ee7ef4.md)
    
- [Abrufen oder Festlegen des Betreffs, wenn Sie einen Termin oder eine Nachricht in Outlook verfassen](654b41f7-97c6-4abd-9eb4-3be337db82f2.md)
    
- [Abrufen oder Festlegen des Orts, wenn Sie einen Termin in Outlook verfassen](23558d28-e6f4-42e2-b943-27354c255a56.md)
    
- [Abrufen oder Festlegen der Uhrzeit, wenn Sie einen Termin in Outlook verfassen](7f3f8159-bf96-4dd4-aedc-e773e36904ea.md)
    
