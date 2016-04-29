
# LabsJS.Labs
Erhalten Sie eine Übersicht der APIs im LabsJS.Labs-Modul.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Das LabsJS.Labs-Modul enthält den Satz von Schlüssel-JavaScript- APIs, die Sie zum Erstellen von Office-Add-Ins (Labs) verwenden können. Die APIs bieten einen Einstiegspunkt für die Lab-Entwicklung.

## LabsJS.Labs API-Modul

Das Labs-Modul enthält die folgenden Typen:


### Variablen


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|Verwenden Sie dieses Objekt, um eine standardmäßige [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)-Instanz zu erstellen.|

### Funktionen


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|Initialisiert eine Verbindung mit dem Host.|
|[Labs.connect (overload)](../../reference/office-mix/labs.connect-overload.md)|Initialisiert eine Verbindung mit dem Host und stellt Eingabeparameter bereit.|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|Initialisiert eine Verbindung mit dem Host.|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|Ruft Konfigurationsinformationen ab, die einer angegebenen Verbindung zugeordnet sind.|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|Trennt das Lab vom Host und stellt einen Lab-Abschlussstatus bereit.|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|Öffnet das angegebene Lab für die Bearbeitung. Sie können die Konfigurationsdaten des Labs angeben, während Sie sich im Bearbeitungsmodus befinden. Ein Lab kann jedoch nicht bearbeitet werden, während es verwendet wird (d. h. während das Lab ausgeführt wird).|
|[Labs.takeLab](../../../reference/office-mix/labs.takelab.md)|Führt das angegeben Lab aus und ermöglicht das Senden von Lab-Ergebnissen an den Server. Beachten Sie, dass Sie ein Lab nicht ausführen können, während es bearbeitet wird.|
|[Labs.on](../../reference/office-mix/labs.on.md)|Fügt einen neuen Handler für ein bestimmtes Ereignis hinzu.|
|[Labs.off](../../reference/office-mix/labs.off.md)|Entfernt einen Ereignishandler für ein angegebenes Ereignis.|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|Ruft eine [Labs.Timeline](../../reference/office-mix/labs.timeline.md)-Objektinstanz ab, die Sie verwenden können, die Sie verwenden können, um Host Player-Steuerelement zu steuern.|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|Deserialisiert ein angegebenes JSON-Objekt in ein Objekt. Sollte nur Komponentenerstellern verwendet werden.|

### Klassen


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md)|Basisklasse für die Initialisierung von Komponenteninstanzen.|
|[Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)|Stellt eine Instanz einer Komponente dar, die eine Instanziierung einer bestimmten Komponente für einen Benutzer zur Laufzeit ist. Das Objekt enthält eine übersetzte Ansicht der Komponente für eine bestimmte Ausführung eines Labs.|
|[Labs.Command](../../reference/office-mix/labs.command.md)|Allgemeiner Befehl, der verwendet wird, um Nachrichten zwischen dem Client und Server zu übergeben.|
|||
|[Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md)|Mit dem  **LabEditor**-Objekt können Sie ein bestimmtes Lab bearbeiten und auch Konfigurationsdaten, die diesem Lab zugeordnet sind, abrufen und festlegen.|
|[Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md)|Eine Instanz eines Labs, das für den aktuellen Benutzer konfiguriert ist. Verwenden Sie dieses Objekt, um Lab-Daten für den Benutzer aufzuzeichnen und abzurufen.|
|||
|||
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|Bietet Zugriff auf die labs.js-Zeitachsenfunktion.|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|Ein Containerobjekt, das Werte für ein angegebenes Lab speichert und verfolgt. Der Wert kann entweder lokal oder auf dem Server gespeichert werden.|

### Schnittstellen


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|Ermöglicht das Abrufen von Daten, die einem [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md)-Befehl zugeordnet sind.|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|Schnittstelle, über die Sie Ereignishandler definieren können.|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|Bietet Möglichkeiten zur Interaktion mit dem [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx)-Objekt.|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|Daten, die einem [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx)-Befehl zugeordnet sind.|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|Daten, die einem Take Action-Befehl zugeordnet sind.|

### Enumerationen


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|Listet mögliche Verbindungzustände des zu hostenden Labs auf.|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)||
