
# Konfigurieren und Bearbeiten von LabsJS-Laboren für Office Mix
Erfahren Sie, wie Sie Ihr Labor konfigurieren müssen, damit das Anzeigen und Bearbeiten unterstützt wird.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Office Mix bietet office.js-Methoden zum Abrufen und Festlegen von Lab-Konfigurationen. Die Konfiguration gibt Office Mix an, welchen Typ Sie erstellen und welchen Datentyp Lab zurückgibt. Diese Informationen dienen der Erfassung und Visualisierung der Analyse.

## Abrufen des Lab-Editors

Mit dem Lab-Editor, dem [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md)-Objekt, können Sie Ihr Lab bearbeiten und Ihre Lab-Konfiguration abrufen und Festlegen. Wenn Sie mit der Bearbeitung Ihres Lab fertig sind, müssen Sie die  **Done**-Methode aufrufen. Das Aufrufen der  **Done**-Methode ist jedoch nur dann erforderlich, wenn Sie versuchen ein Lab zu verwenden, das Sie bearbeiten. Beachten Sie, dass nur eine Instanz des Lab gleichzeitig geöffnet werden kann.

Der folgende Code veranschaulicht, wie der Lab-Editor abgerufen werden kann.




```
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Verwenden Sie die  **getConfiguration**- und  **setConfiguration**-Methoden zum [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md), um die Konfiguration eines angegebenen Labs zu speichern. Die Konfiguration ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) gibt Office Mix an, welche Daten von Lab erfasst und verarbeitet werden. Eine Konfiguration enthält allgemeine Informationen zu einem Lab, dazu gehört der Name, die Version und sonstige Konfigurationsoptionen. Der wichtigste Teil der Konfiguration ist die Definition der Lab-Komponente.

Der folgende Code veranschaulicht, wie eine Konfiguration abgerufen und festgelegt werden kann. Um eine Konfiguration festzulegen, müssen Sie das Konfigurationsobjekt erstellen und anschließend die  **setConfiguration**-Methode aufrufen. Rufen Sie dann die  **getConfiguration**-Methode zu dem Lab Editor-Objekt auf, um die Konfiguration abzurufen.




```

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## Schließen des Editors

Rufen Sie die  **Done**-Methode zum Editor auf, wenn Sie mit dem Bearbeiten des Lab fertig sind, um den Editor zu schließen. Beachten Sie, dass Sie ein Lab nicht gleichzeitig verwenden und bearbeiten können. Nachdem Sie  **Done** aufgerufen haben, können Sie das Lab jedoch bearbeiten oder ausführen.


## Interagieren mit einem Lab

Nachdem Sie die Lab-Konfiguration festgelegt haben, können Sie mit dem Interagieren mit dem Lab beginnen. Wenn das Lab in PowerPoint ausgeführt wird, werden Interaktionen simuliert. Wenn das Lab im Office Mix Lesson Player ausgeführt wird, werden die Daten jedoch in der Office Mix-Datenbank gespeichert und in der Analyse verwendet.


### Abrufen der Lab-Instanz

Sie interagieren mit Lab mithilfe des [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md)-Objekts, dabei handelt es sich um eine Instanz des konfigurierten Lab für den aktuellen Benutzer. Rufen Sie die [Labs.takeLab](../../../reference/office-mix/labs.takelab.md)-Funktion auf, um das Lab auszuführen (zu verwenden).


```
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

Das Instanzobjekt enthält ein Array von Komponenteninstanzen ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md), [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)), die den von Ihnen in der Konfiguration angegebenen Komponenten zugeordnet werden. Eine Instanz ist einfach eine transformierte Version der Konfiguration, die zum Anfügen der serverseitigen IDs an Instanzobjekte verwendet wird, und zum Ausblenden bestimmter Felder vom Benutzer, falls zutreffend (z. B. Tipps, Antworten usw.).


### Verwaltungszustand

Zustand ist der temporäre mit dem Benutzer verknüpfte Speicher, der ein bestimmtes Lab ausführt. Sie können den Speicher verwenden, um Informationen zwischen aufeinanderfolgenden Aufrufen des Lab beizubehalten. Ein Programmierungslab könnte beispielsweise die aktuelle Arbeit des Benutzers speichern.

Verwenden Sie folgenden Code, um  **set** anzugeben.




```
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

Verwenden Sie folgenden Code, um  **get** anzugeben.




```
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## Komponenteninstanzen und Ergebnisse

Im Folgenden finden Sie eine Übersicht dazu, wie die Instanzen der vier Komponententypen implementiert werden, und kurze Beispiele der Komponentenmethoden. 

Zunächst müssen Sie sich jedoch mit den beiden Kernkonzepten beim Arbeiten mit Komponenteninstanzen vertraut machen. Das erste ist das Konzept der  **Versuche** und **Werte**.

 **Versuche**

Ein Versuch ist ein Versuch von einem Benutzer zum Abschluss einer Komponenteninstanz. Im Fall einer Multiple-Choice-Frage startet ein Versuch, wenn der Benutzer an der Lösung des Problems arbeitet und endet, wenn eine endgültige Bewertung zugewiesen wird. Die Office Mix-Analyse aggregiert anschließend die Benutzerergebnisse für das Problem.


 >**Hinweis**  Versuche können für alle Komponententypen außer für den  **DynamicComponent**-Typ verwendet werden.

Sie können die Ergebnisse für alle auf eine bestimmte Komponenteninstanz bezogenen Versuche abrufen, indem Sie die  **getAttempts**-Methode verwenden. Nachdem Sie die Ergebnisse abgerufen haben, kann der Benutzer es entweder erneut versuchen, einen der vorhandenen Versuche zu verwenden, indem er die  **resume**-Methode verwendet, oder einen neuen Versuch erstellt, indem er die  **createAttempt**-Methode verwendet. Das folgende Beispiel veranschaulicht den Prozess.




```
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Werte**

Komponenteninstanzen enthalten ein Wörterbuch mit Schlüsseln, die einem Array von Werten zugeordnet sind. Sie können das Array verwenden, um Tipps, Feedback oder andere Wertesätze zu speichern, die Sie mit der Komponente verknüpfen möchten. Die Komponenteninstanz bietet Zugriff auf diese Werte mithilfe der  **getValues**-Methode. 

Durch das Abfragen eines Tipp-Werts kann die Analyse markieren, dass der Benutzer einen Tipp genutzt hat. Die Werte werden auf pro-Versuch-Basis nachverfolgt.

Im folgenden Codebeispiel wird veranschaulicht, wie ein Tipp abgefragt wird.




```
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### ActivityComponentInstance


Verwenden Sie das  **ActivityComponentInstace**-Objekt, um die Benutzeraktionen mit einer Aktivitätskomponente nachzuverfolgen. Diese Klasse bietet eine  **complete**-Methode, um anzugeben, dass der Benutzer die Interaktion mit der Aktivität abzuschließen. Die Methode kann angeben, dass der Benutzer eine zugewiesene Aufgabe abzuschließen, ein Lesevorgang abgeschlossen hat, oder ein anderer der Aktivität zugewiesener Endpunkt erreicht wurde. Der folgende Code veranschaulicht, wie die  **complete**-Methode verwendet wird.


```
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### ChoiceComponentInstance


Verwenden Sie das  **ChoiceComponentInstance**-Objekt, um die Benutzerinteraktionen mit einer Auswahlkomponente nachzuverfolgen. Auswahlkomponenten sind Probleme, die dem Benutzer eine Liste von Auswahloptionen bereitstellen, aus denen er eine auswählen muss. Dabei kann eine korrekte Antwort geben, dem muss aber nicht so sein. Die Klasse stellt zwei primäre Methoden bereit:  **getSubmissions** und **submit**. Mit der  **getSubmissions**-Methode können Sie zuvor gespeicherte Übermittlungen abrufen; mit der  **submit**-Methode kann eine neue Übermittlung gespeichert werden. Die folgenden Codebeispiele veranschaulichen die Verwendung der Methoden.


```
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### InputComponentInstance


Verwenden Sie das  **InputComponentInstance**-Objekt, um die Benutzeraktionen mit einer Eingabekomponente nachzuverfolgen. Die Klasse bietet zwei primäre Methoden:  **getSubmission** und **submit**. Mit der  **getSubmissions**-Methode können Sie zuvor gespeicherte Übermittlungen abrufen; mit der  **submit**-Methode können Sie eine neue Übermittlung speichern. Der folgende Codeausschnitt veranschaulicht die Verwendung der  **getSubmissions**-Methode.


```
var submissions = this._attempt.getSubmissions();
```

Beachten Sie bei der Verwendung der  **submit**-Methode, dass das  **InputComponentAnswer**-Objekt die übermittelte Antwort darstellt und das  **InputComponentResult**-Objekt das Ergebnis enthält. Der Rückgabewert ist ein  **InputComponentSubmission**-Objekt, das die Antwort, das Ergebnis und einen Zeitstempel enthält, der angibt, dass das Ergebnis übermittelt wurde.




```
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### DynamicComponentInstance


Verwenden Sie das  **DynamicComponentInstance**-Objekt, um die Benutzerinteraktionen mit einer dynamischen Komponente nachzuverfolgen. Die primären Methoden in dieser Klasse sind  **getComponents**,  **createComponent** und **close**.

Mit der  **getComponents**-Methode können Sie eine Liste der zuvor erstellten Komponenteninstanz abrufen, siehe folgendes Beispiel.




```
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

Die  **createComponent**-Methode erstellt eine neue Komponente und gibt die Komponenteninstanz zurück, siehe folgendes Beispiel.




```
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Verwenden Sie die  **close**-Methode, um anzugeben, dass Sie die Verwendung der dynamischen Komponente abgeschlossen haben, um neue Komponenten zu erstellen. Beachten Sie, dass Sie auch eine  **isClosed** boolesche Methode verwenden können, um zu testen, ob die dynamische Komponenteninstanz geschlossen wurde. Der folgende Code veranschaulicht, wie die **close**-Methode verwendet wird.




```
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## Zusätzliche Ressourcen



- [Office Mix-Add-Ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Exemplarische Vorgehensweise: Erstellen der ersten Übungseinheit für Office Mix](walkthrough:-creating-your-first-lab-for-office-mix.md)
    
