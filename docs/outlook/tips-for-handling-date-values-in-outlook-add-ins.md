
# Tipps zum Verwenden von Datumswerten in Outlook-Add-Ins
Erfahren Sie Tipps dazu, wie Outlook-Add-Ins Werte für das Eingangs- und Ausgangsdatum für die unterschiedlichen datumsbezogenen Methoden in der JavaScript-API für Office handhaben sollen.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_


## Allgemeine JavaScript-Unterstützung für Datum und Uhrzeit


Die JavaScript-API für Office verwendet für die meisten Aktionen zum Speichern und Abrufen von Datum und Uhrzeit ein JavaScript- [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp)-Objekt. Dieses  **Date**-Objekt stellt Methoden wie [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) und [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp) bereit, die den angeforderten Datums- oder Uhrzeitwert entsprechend der Universal Coordinated Time (UTC, koordinierte Weltzeit) zurückgeben.

Das  **Date**-Objekt stellt auch andere Methoden bereit wie [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) und [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), die die angeforderten Datums- oder Uhrzeitwerte entsprechend der Ortszeit zurückgeben.

Die "Ortszeit" ist im Wesentlichen vom Browser und dem Betriebssystem auf dem Clientcomputer abhängig. Beispielsweise gibt ein JavaScript-Aufruf von  **getDate** auf einem Windows-Computer in den meisten Fällen ein Datum basierend auf der auf dem Clientcomputer in Windows eingestellten Zeitzone zurück.

Das folgende Beispiel erstellt ein  **Date**-Objekt  `myLocalDate` in der Ortszeit und ruft **toUTCString** auf, um dieses Datum in eine Datumszeichenfolge in UTC zu konvertieren.




```
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Sie können zwar das JavaScript-Objekt  **Date** verwenden, um einen Datums- oder Zeitwert auf der Grundlage von UTC oder der Zeitzone des Clientcomputers abzurufen, das **Date**-Objekt besitzt jedoch in einer Hinsicht eine Einschränkung: Es stellt keine Methoden zur Rückgabe von Datums- oder Zeitwerten für andere spezifische Zeitzonen bereit. Wenn für Ihren Clientcomputer beispielsweise Eastern Standard Time (EST) eingestellt ist, steht keine  **Date**-Methode zur Verfügung, den Stundenwert einer anderen Zeitzone als EST oder UTC, beispielsweise Pacific Standard Time (PST), abzurufen. 


## Datumsspezifische Funktionen für Outlook-Add-In


Die vorgenannten JavaScript-Einschränkungen sind für Sie von Bedeutung, wenn Sie die JavaScript-API für Office zur Verarbeitung von Datums- oder Zeitwerten in Outlook-Add-Ins verwenden, die auf einem Outlook Rich Client und in Outlook Web App oder OWA für mobile Geräte ausgeführt werden.


### Zeitzonen für Outlook-Clients

Lassen Sie uns zur Verdeutlichung die fraglichen Zeitzonen definieren.



|**Zeitzone**|**Beschreibung**|
|:-----|:-----|
|Zeitzone des Clientcomputers|Diese wird im Betriebssystem des Clientcomputers festgelegt. Die meisten Browser verwenden die Zeitzone des Clientcomputers, um Datums- oder Zeitwerte des JavaScript-Objekts  **Date** anzuzeigen.Ein Outlook Rich Client verwendet diese Zeitzone, um Datums- oder Zeitwerte auf der Benutzeroberfläche anzuzeigen. Beispielsweise verwendet Outlook auf einem Clientcomputer unter Windows die in Windows festgelegte Zeitzone als lokale Zeitzone. Wenn der Benutzer die Zeitzone auf einem Mac-Clientcomputer ändert, fordert Outlook für Mac den Benutzer auf, die Zeitzone auch in Outlook zu aktualisieren.|
|EAC-Zeitzone (Exchange Admin Center)|Dieser Zeitzonenwert (und die bevorzugte Sprache) wird vom Benutzer festgelegt, wenn er sich erstmals bei Outlook Web App oder OWA für mobile Geräte anmeldet. Outlook Web App und OWA für mobile Geräte verwenden diese Zeitzone, um Datums- oder Zeitwerte auf der Benutzeroberfläche anzuzeigen.|
Da ein Outlook Rich Client die Zeitzone des Clientcomputers verwendet und die Benutzeroberfläche von Outlook Web App und OWA für mobile Geräte die EAC-Zeitzone, kann die lokale Zeitzone für ein für das gleiche Postfach installierte Outlook-Add-In bei Ausführung auf einem Outlook Rich Client und in Outlook Web App oder OWA für mobile Geräte unterschiedlich sein. Als Entwickler von Mail-Add-ins sollten Sie Datumswerte entsprechend ein- und ausgeben, damit diese Werte immer konsistent mit der Zeitzone sind, die der Benutzer auf dem entsprechenden Client erwartet.


### Datumsspezifische API

Die folgenden Eigenschaften und Methoden in der JavaScript-API für Office unterstützen die datumsspezifischen Features.



|**API-Element**|**Zeitzonendarstellung**|**Beispiel unter einem Outlook Rich Client**|**Beispiel in Outlook Web App oder OWA für mobile Geräte**|
|:-----|:-----|:-----|:-----|
|[Office.context.mailbox.userProfile.timeZone](../reference/outlook/Office.context.mailbox.userProfile.md%28Office.15%29.md)|Unter einem Outlook Rich Client gibt diese Eigenschaft die Zeitzone des Clientcomputers zurück. Unter Outlook Web App und OWA für mobile Geräte gibt diese Eigenschaft die Zeitzone EAC zurück. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/outlook/Office.context.mailbox.item.html%28Office.15%29.md) und [Office.context.mailbox.item.dateTimeModified](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md%28Office.15%29.md)|Jede dieser Eigenschaften gibt ein JavaScript-Objekt  **Date** zurück. Dieser **Date**-Wert ist UTC-richtig, wie im folgenden Beispiel dargestellt.  `myUTCDate` besitzt den gleichen Wert in einem Outlook Rich Client, in Outlook Web App und in OWA für mobile Geräte.
```
var myDate = Office.mailbox.item.dateTimeCreated;
var myUTCDate = myDate.getUTCDate;
```

Durch den Aufruf von  `myDate.getDate` wird jedoch ein Datumswert entsprechend der Zeitzone des Clientcomputers zurückgegeben, die derjenigen entspricht, die zur Anzeige von Datums- und Zeitwerten auf der Benutzeroberfläche des Outlook Rich Client verwendet wird, die sich jedoch von der EAC-Zeitzone unterscheiden kann, die Outlook Web App und OWA für mobile Geräte auf seiner Benutzeroberfläche verwenden.|Wenn das Element um 9 Uhr UTC erstellt wird: Gibt  `Office.mailbox.item.dateTimeCreated.getHours` 4 Uhr EST zurück.Wenn das Element um 11 Uhr UTC geändert wird: Gibt  `Office.mailbox.item.dateTimeModified.getHours` 6 Uhr EST zurück.|Wenn die Erstellungszeit 9 Uhr UTC ist: Gibt  `Office.mailbox.item.dateTimeCreated.getHours` 4 Uhr EST zurück.Wenn das Element um 11 Uhr UTC geändert wird: Gibt  `Office.mailbox.item.dateTimeModified.getHours` 6 Uhr EST zurück.Beachten Sie bei der Anzeige der Erstellungs- und Änderungszeiten auf der Benutzeroberfläche, dass Sie diese zunächst in PST konvertieren sollten, damit diese konsistent mit dem Rest der Benutzeroberfläche sind.|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/outlook/Office.context.mailbox.md%28Office.15%29.md)|Sowohl der  _Start_- als auch der  _End_-Parameter benötigt ein JavaScript-Objekt  **Date**. Die Argumente sollten unabhängig von der auf der Benutzeroberfläche von einem Outlook Rich Client, von Outlook Web App oder OWA für mobile Geräte verwendeten Zeitzone UTC-richtig sein. |Wenn die Anfangs- und Endzeiten des Terminformulars 9 Uhr UTC und 11 Uhr UTC angeben, sollten Sie sicherstellen, dass das  `start`- und  `end`-Argument UTC-richtig sind, das bedeutet: 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">start.getUTCHours</span> gibt 9 Uhr UTC zurück.</p></li><li><p><span class="code">end.getUTCHours</span> gibt 11 Uhr UTC zurück.</p></li></ul>|Wenn die Anfangs- und Endzeiten des Terminformulars 9 Uhr UTC und 11 Uhr UTC angeben, sollten Sie sicherstellen, dass das  `start`- und  `end`-Argument UTC-richtig sind, das bedeutet: 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><span class="code">start.getUTCHours</span> gibt 9 Uhr UTC zurück.</p></li><li><p><span class="code">end.getUTCHours</span> gibt 11 Uhr UTC zurück.</p></li></ul>|

## Hilfsmethoden für datumsspezifische Szenarien


Da sich die Ortszeit, wie in den vorherigen Abschnitten beschrieben, für Benutzer unter Outlook Web App bzw. OWA für mobile Geräte auf einem Outlook Rich Client unterscheiden kann, das JavaScript-Objekt  **Date** jedoch nur die Konvertierung in die Zeitzone des Clientcomputers oder in UTC unterstützt, verfügt die JavaScript-API für Office über zwei Hilfsmethoden: [Office.context.mailbox.convertToLocalClientTime](../reference/outlook/Office.context.mailbox.html%28Office.15%29.md) und [Office.context.mailbox.convertToUtcClientTime](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.md%28Office.15%29.md). 

Diese Hilfsmethoden verarbeiten alle Datums- und Zeitwerte, die für die folgenden beiden datumsspezifischen Szenarien in einem Outlook Rich Client, Outlook Web App und OWA für mobile Geräte unterschiedlich gehandhabt werden müssen, sodass das Konzept „einmal schreiben" für unterschiedliche Clients Ihres Outlook-Add-Ins unterstützt wird.


### Szenario A: Anzeigen der Erstellungs- oder Änderungszeit des Elements

Wenn Sie die Erstellungszeit ( **Item.dateTimeCreated**) oder die Änderungszeit ( **Item.dateTimeModified**) des Elements auf der Benutzeroberfläche anzeigen, verwenden Sie zunächst  **convertToLocalClientTime**, um das von diesen Eigenschaften zum Abrufen einer Wörterbuchdarstellung bereitgestellte  **Date**-Objekt in die entsprechende Ortszeit zu konvertieren. Das Folgende ist ein Beispiel dieses Szenarios:


```
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Beachten Sie, dass  **convertToLocalClientTime** die Unterschiede zwischen einem Outlook Rich Client und Outlook Web App bzw. OWA für mobile Geräte verarbeitet:


- Wenn  **convertToLocalClientTime** feststellt, dass der aktuelle Host ein Rich Client ist, konvertiert die Methode die **Date**-Darstellung in eine Wörterbuchdarstellung in der gleichen Clientcomputer-Zeitzone, die konsistent mit der restlichen Benutzeroberfläche des Rich Client ist.
    
- Wenn  **convertToLocalClientTime** feststellt, dass der aktuelle Host Outlook Web App ist, konvertiert die Methode die UTC-richtige **Date**-Darstellung in ein Wörterbuchformat in der Zeitzone EAC, die konsistent mit dem Rest der Outlook Web App-Benutzeroberfläche ist.
    

### Szenario B: Anzeigen von Anfangs- und Enddatum in einem neuen Terminformular

Wenn Sie als Eingabe unterschiedliche Teile eines in Ortszeit dargestellten Datums- und Zeitwerts erhalten und diesen Wörterbuch-Eingabewert als Anfangs- oder Endzeit in einem Terminformular verwenden möchten, verwenden Sie zunächst die  **convertToUtcClientTime**-Hilfsmethode, um den Wörterbuchwert in ein UTC-richtiges  **Date**-Objekt zu konvertieren. 

Im folgenden Beispiel wird davon ausgegangen, dass es sich bei  `myLocalDictionaryStartDate` und `myLocalDictionaryEndDate` um Datums- und Zeitwerte im Wörterbuchformat handelt, die Sie vom Benutzer erhalten haben. Diese Werte basieren auf der von der Hostanwendung abhängigen Ortszeit.




```
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Die resultierenden Werte  `myUTCCorrectStartDate` und `myUTCCorrectEndDate`, sind UTC-richtig. Übergeben Sie anschließend diese  **Date**-Objekte als Argumente für den  _Start_- und  _End_-Parameter der  **Mailbox.displayNewAppointmentForm**-Methode, um das neue Terminformular anzuzeigen.

Beachten Sie, dass  **convertToUtcClientTime** die Unterschiede zwischen einem Outlook Rich Client und Outlook Web App bzw. OWA für mobile Geräte verarbeitet:


- Wenn  **convertToUtcClientTime** feststellt, dass es sich bei dem aktuellen Host um einen Outlook Rich Client handelt, konvertiert die Methode die Wörterbuchdarstellung einfach in ein **Date**-Objekt. Dieses  **Date**-Objekt ist UTC-richtig, wie von  **displayNewAppointmentForm** erwartet.
    
- Wenn  **convertToUtcClientTime** feststellt, dass es sich bei dem aktuellen Host um Outlook Web App handelt, konvertiert die Methode das Wörterbuchformat der in der Zeitzone EAC angezeigten Datums- und Zeitwerte in ein **Date**-Objekt. Dieses  **Date**-Objekt ist UTC-richtig, wie von  **displayNewAppointmentForm** erwartet.
    

## Zusätzliche Ressourcen



- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](d6eea4c4-bb21-4f24-bcba-1eccbb4e12dd.md)
    


