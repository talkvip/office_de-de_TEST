
# Richtlinien für die Erstellung von Laboren für Mix mit LabsJS
Erhalten Sie eine präzise Einführung in die Verwendung der LabsJS JavaScript-Bibliothek, um Labore für Office Mix zu erstellen.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Die LabsJS-Bibliothek (labs.js) unterstützt das Schreiben spezieller Office-Add-Ins (aufgerufene Labore), die in Office Mix. Office Mix integriert wird, anschließend werden die Labore mit Microsoft PowerPoint gerendert. Obwohl wir diese Komponenten „Labs" nennen, möchten wir klarstellen, dass wir spezielle Office-Add-Ins erstellen, die Office Mix-Add-ins sind.

Die LabsJS-Inhalte helfen Ihnen bei der Implementierung der labs.js JavaScript-API durch die Bereitstellung von Anleitungen und Beispielen. Diese Bibliothek setzt auf die [JavaScript-API für Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx) (Office.js) auf und bietet eine Abstraktionsebene, die für in Office Mixeingebettete Add-Ins optimiert ist.


## Allgemeine Richtlinien


Im Folgenden finden Sie einige allgemeine Richtlinien, die Sie beim Schreiben von Add-Ins mithilfe der LabJS-API unterstützen.


### Skripts

Da die labs.js-Bibliothek eine Abstraktionsebene der office.js ist und deshalb eine Abhängigkeit zu office.js herrscht,müssen die office.js- und labs.js-Bibliotheksdateien in Ihre Entwicklungsprojekte miteinbezogen werden. 

Sie können folgendermaßen auf die office.js-Bibliothek verweisen:  `<script src="https://sforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>`.

Die labs.js-Bibliothek ist im LabsJS SDK enthalten. Alternativ können Sie im CDN unter  `https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js` auf die labs.js-Bibliothek verweisen. Beachten Sie, dass die Produktionsversion Ihres Labors auf die im CDN gespeicherte Version verweisen muss.


 >**Hinweis**  Zusätzlich zu der JavaScript-Datei (labs-1.0.4.js) bieten wir eine TypeScript-Definitionsdatei der Labs-API (labs-1.0.4.d.ts). Die Definitionsdatei wurde mit TypeScript Version 0.9.1.1 erstellt.


### Rückrufe und Fehlerbehandlung

Verschiedene Methoden in der labs.js-API arbeiten asynchron. Für diese Vorgänge übernimmt die API eine Standard-Rückrufoberfläche an,  **ILabCallback**. 


```
function(err, result) {
}
```

Die Rückrufmethode verwendet zwei Parameter,  _err_ und _result_. Das  _err_-Feld bleibt  **null**, sofern kein Fehler auftritt. Das  _result_-Feld gibt das Ergebnis des Vorgangs zurück.

Der Rückrufvorgang wird niemals sofort ausgelöst, selbst wenn das Ergebnis sofort verfügbar ist. Stattdessen wird sie bei einer separaten Ausführung der JavaScript-Ereignisschleife (mithilfe des  **setTimeout**-Aufrufs) ausgelöst. Durch das Übernehmen dieser Rückrufdefinition können Sie labs.js leicht in Ihre Wunsch-API integrieren. Sie können beispielsweise jQuery-Zusagen für diese Rückrufe mit einer einfachen Übersetzungsmethode austauschen, siehe folgendes Beispiel.




```
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### Lab-Host und DefaultLabHost

Das Lab-Host ( **ILabHost**) ist der zugrundeliegende Treiber, der die Entwicklung von Labs unterstützt. Standardmäßig wird das auf einen Host festgelegt, der in office.js integriert ist.

Zu Testzwecken und zum Ausführen Ihres Labors innerhalb von labhost.html müssen Sie zu einem Host wechseln, dass in der Simulationsumgebung funktioniert. Im folgenden Codebeispiel wird aufgezeigt, wie Sie dazu mit einem Abfrageparameter vorgehen müssen. Alternativ können Sie  **DefaultHostBuilder** so ändern, dass er mit einer anderen Plattform in Ihr Lab-Add-In integriert werden kann.




```

Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### Initialisierung

Die Initialisierung schafft das Kommunikationsumfeld zwischen dem Labor und dem Host. Initialisieren Sie Ihr Labor mit folgendem Aufruf.


```
Labs.connect((err, connectionResponse) => {});
```

Nach der Initialisierung können Sie andere Methoden der labs.js-API aufrufen. Der  _connectionResponse_-Parameter enthält Informationen zum Host, Benutzer und anderen verbindungsrelevanten Informationen. Weitere Informationen zu den zurückgegebenen Werten finden Sie unter [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md).


### Zeitformat

Labs.js speichert Zahlen als Millisekunden, die seit dem 1. Januar 1970 UTC vergangen sind. Dies entspricht dem Datumsformat des JavaScript [Date-Objekts](http://msdn.microsoft.com/de-de/library/ie/cd9w2te4%28v=vs.94%29.aspx),


### Zeitachse

Das Labor kann außerdem mit der Lesson Player-Zeitachse interagieren. Mit der Zeitachse kann das Labor dem Lesson Player mitteilen, mit der nächsten Folie fortzufahren. Das Zeitachsenobjekt wird abgerufen, indem die  **Labs.getTimeline**-Methode aufgerufen wird.


```
Labs.getTimeline().next({}, (err, unused) => { });
```


## Behandeln von Ereignissen


Die LabsJS-Ereignis-API verfolgt laborrelevante Ereignisse und ermöglicht Ihnen, Ereignishandler hinzuzufügen, sodass Sie auf Ereignisse reagieren oder darin agieren. Die Ereignismethoden, von denen drei vorhanden sind, befinden sich im  **EventTypes**-Objekt:  **ModeChanged**,  **Activate**, and  **Deactivate**. 


### Modusänderung

Das  **ModeChanged**-Ereignis wird ausgelöst, wenn das angegebenen Labor zwischen Bearbeitungsmodus und Anzeigemodus geändert wird. Der Bearbeitungsmodus wird angezeigt, wenn das Labor im PowerPoint-Bearbeitungsmodus angezeigt wird. Der Anzeigemodus wird angezeigt, wenn PowerPoint die Bildschirmpräsentation rendert oder wenn das Labor im Office Mix Lesson Player angezeigt wird. Der Anzeigemodus zeigt immer an, was der Benutzer sieht, wenn er das Labor verwendet. Mit dem Bearbeitungsmodus kann der Benutzer das Labor konfigurieren. 

Daten im  **ModeChangedEventData**-Objekt, das an den Rückruf weitergegeben wird, enthält Informationen zum aktuellen Modus. Der folgende Code zeigt, wie das  **ModeChanged**-Ereignis verwendet wird.




```
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### Aktivieren

Das  **activate**-Ereignis wird ausgelöst, wenn die PowerPoint-Folie im Lesson Player aktiviert wird, auf der sich das Labor befindet.


```
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### Deaktivieren

Das  **deactivate**-Ereignis wird ausgelöst, wenn die PowerPoint-Folie, auf der sich das Labor befindet, nicht mehr die aktive Folie ist.


```
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### Zeitachse

Das Labor kann auch mit der Lesson Player-Zeitachse interagieren. Mit der Zeitachse kann das Labor dem Lesson Player mitteilen, dass er mit der nächsten Folie fortfahren soll. Das Zeitachsenobjekt wird durch Aufrufen der  **Labs.getTimeline**-Methode abgerufen.


```
Labs.getTimeline().next({}, (err, unused) => { });
```


## Zusätzliche Ressourcen



- [Office Mix-Add-Ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
