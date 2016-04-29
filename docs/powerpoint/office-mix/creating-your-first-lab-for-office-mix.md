
# Exemplarische Vorgehensweise: Erstellen der ersten Übungseinheit für Office Mix
Erstellen Sie Ihr erstes LabsJS-Labor mithilfe der Schritt-für-Schritt-Anleitung.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

In dieser exemplarischen Vorgehensweise erstellen Sie ein einfaches LabsJS-Labor von Grund auf neu. Ihr Labor ist ein einfaches wahr/falsch-Quiz, das eine einzelne Frage stellt. 

Statt mit einer Visual Studio-Projektvorlage zu beginnen, starten Sie mit drei leeren Dateien – dies zeigt, wie einfach eine Übungseinheit ist: 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
Sie können einen beliebigen Code-Editor verwenden, um diese Dateien zu bearbeiten, da wir nicht mit einer Visual Studio-Vorlage beginnen. Die HTML-Datei ist im Grunde genommen trivial und wenn Sie möchten, können Sie das HTML-Markup einfach per "Copy und Paste" aus den Lernprogramm-Dateien extrahieren. Beachten Sie jedoch, dass es sich dabei um HTML5 handeln muss, vergewissern Sie sich also, dass Ihre doctype-Deklaration  `<!DOCTYPE html>` lautet. Die CSS-Datei ist optional. Der schwierige Teil erfolgt in der JavaScript-Datei (.js), TrueFalse.js.
In der exemplarischen Vorgehensweise werden die vier wichtigsten Labor-Features behandelt:

- Einrichtung (Herstellen einer Verbindung mit dem Host)
    
- Modusänderungen (zwischen Bearbeitungsmodus und Anzeigemodus)
    
- Bearbeiten des Labors
    
- Erstellen (oder Ausführen) der Labor
    

 >**Hinweis**  Die Datei labhost.html wird auf einem Webserver ausgeführt und stellt die Hosting-Umgebung für die Labor-Entwicklung und das Testing bereit. Dadurch wird die Labor-Entwicklung erheblich vereinfacht. Weitere Informationen zur Einrichtung Ihrer Entwicklungsumgebung finden Sie unter [Erste Schritte mit LabsJS für Office Mix](../../powerpoint/office-mix/get-started-with-labsjs-for-office-mix.md).

Schließlich wird Ihnen die fertig gestellte JavaScript-Datei (TrueFalse.js) mit den anderen mit diesem SDK bereitgestellten Dateien angezeigt. Darauf folgt eine Anleitung zum Programmierprozess.

## Herstellen einer Verbindung mit dem Labor-Host

Labore in dieser Umgebungen können mit unserem Labor-Host (für Entwicklung und Testing) oder mit dem Standard-Runtimehost, das vom Office.js-Host bereitgestellt wird. Die Funktion zum Öffnen verwendet anschließend einen einfachen if/else-Ausdruck, um zu testen, welche von diesen Hostingkontexten zutrifft.


```
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

Das  **PostMessageLabHost**-Objekt wird in der labhost.html-Entwicklungsumgebung, in der Produktion wird das Labor mithilfe des  **OfficeJSLabHost** in PowerPoint/Office Mix ausgeführt.

Erstellen Sie als Nächstes eine Hilfsmethode, um einen Rückruf zu erstellen, dessen Aufgabe darin besteht, ein von Ihnen eingereichtes zurückgestelltes jQuery-Objekt aufzulösen oder abzulehnen. Verwenden Sie diese Methode,  **createCallback**, um vom jQuery-Bereich zu den von der labs.js definierten Rückrufen zurückzukehren.




```
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

Außerdem erstellen wir eine Hilfsmethode für die Labor-Konfiguration zum Abrufen einer bestimmten Frage und Antwort.




```
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## Modusänderungen

Ein Labor befindet sich immer in einem von zwei Zuständen oder Modi:  **view** und **edit**. Deshalb müssen wir eine Möglichkeit finden, den Status und das Verhalten für das Quiz zu erfassen und beizubehalten; zu diesem Zweck erstellen wir eine Klasse.


```
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

Außerdem stellen wir eine Hilfsmethode bereit, dessen Aufgabe das Aktualisieren der UI des Quiz ist, auf Grundlage der Antwort auf die Quizfrage (d. h. der „Beitrag") ist korrekt oder falsch.




```
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

Wir benötigen außerdem eine Funktion zum Wechseln zwischen dem Bearbeitungs- und Anzeigemodus.




```
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

Unsere nächste Funktion aktualisiert die Konfiguration für das Quiz auf Grundlage der Änderungsereignisse, die wir von der UI erhalten haben.




```
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

Als Nächstes aktualisieren wir die auf dem Server gespeicherten Laborkonfigurationsdaten, die auf bestimmten Fragen und Antworten basieren.




```
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

Als Nächstes haben wir eine Funktion, die im Labor vorgenommene Aktualisierungen im Bearbeitungsmodus an von uns vorgenommene Konfigurationsänderungen bindet. Im Folgenden finden Sie den Code zum Aufheben der Bindung von den zuvor gebundenen Änderungshandlern.




```
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

Nun kommt der wichtigste Teil des Abschnitts, d. h. Methoden für das Wechseln zwischen dem Anzeige- und Bearbeitungsmodus. Beginnen wir mit dem Wechsel vom Anzeige- in den Bearbeitungsmodus.




```
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

Wechseln wir nun vom Bearbeitungs- in den Anzeigemodus.




```
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

Zum Abschluss, nachdem Sie eine Verbindung zum Host hergestellt haben und das Dokument bereit ist, starten Sie das Quiz.




```
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## Zusätzliche Ressourcen



- [Office Mix-Add-Ins](../../powerpoint/office-mix/office-mix-add-ins.md)
    
