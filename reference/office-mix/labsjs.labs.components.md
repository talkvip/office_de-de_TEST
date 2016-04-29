
# LabsJS.Labs.Components
Bietet eine allgemeine Übersicht über die JavaScript-API von Labs.JS Labs.Components.

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

APIs im Modul "Labs.Components" stellen die vier standardmäßigen Komponenten dar, die derzeit für die Entwicklung von Labs verfügbar sind (Aktivität, Auswahl, Eingabe und Dynamisch).

## Labs.Components-Modul

Nachfolgend sind Labs.Components-Typen aufgeführt:


### Klassen


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](../../reference/office-mix/labs.components.componentattempt.md)|Basisklasse für Versuche mit Komponenten.|
|[Labs.Components.ActivityComponentAttempt](../../reference/office-mix/labs.components.activitycomponentattempt.md)|Stellt einen Versuch zum Abschließen einer Aktivitätskomponente dar.|
|[Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md)|Stellt die aktuelle Instanz einer Aktivitätskomponente dar.|
|[Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md)|Die Antwort auf ein Problem in einer Auswahlkomponente.|
|[Labs.Components.ChoiceComponentAttempt](../../reference/office-mix/labs.components.choicecomponentattempt.md)|Stellt einen Versuch in einer Auswahlkomponente dar.|
|[Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md)|Stellt eine Instanz einer Auswahlkomponente dar.|
|[Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md)|Das Ergebnis der Übermittlung einer Auswahlkomponente.|
|[Labs.Components.ChoiceComponentSubmission](../../reference/office-mix/labs.components.choicecomponentsubmission.md)|Stellt die Übermittlung im Zusammenhang mit einer Auswahlkomponente dar.|
|[Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md)|Stellt eine Instanz einer dynamischen Komponente dar.|
|[Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)|Stellt die Antwort auf ein Problem bei einer Eingabekomponente dar.|
|[Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)|Stellt einen Versuch der Interaktion mit einer Eingabekomponente dar.|
|[Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md)|Stellt eine Instanz einer Eingabekomponente dar.|
|[Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)|Das Ergebnis der Übermittlung einer Eingabekomponente.|
|[Labs.Components.InputComponentSubmission](../../reference/office-mix/labs.components.inputcomponentsubmission.md)|Stellt eine Übermittlung an eine Eingabekomponente dar.|

### Schnittstellen


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](../../reference/office-mix/labs.components.iactivitycomponent.md)|Stellt eine Aktivitätskomponente dar. Erweitert [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md).|
|[Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|Stellt eine bestimmte Instanz einer Aktivitätskomponente dar. Erweitert [Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md).|
|[Labs.Components.IChoice](../../reference/office-mix/labs.components.ichoice.md)|Eine verfügbare Auswahl für ein bestimmtes Problem.|
|[Labs.Components.IChoiceComponent](../../reference/office-mix/labs.components.ichoicecomponent.md)|Ermöglicht Interaktion mit einer Auswahlkomponente.|
|[Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)|Eine Instanz einer Auswahlkomponente.|
|[Labs.Components.IDynamicComponent](../../reference/office-mix/labs.components.idynamiccomponent.md)|Ermöglicht die Interaktion mit einer dynamischen Komponente.|
|[Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md)|Eine Instanz einer dynamischen Komponente.|
|[Labs.Components.IHint](../../reference/office-mix/labs.components.ihint.md)|Hinweis bei einem Lab-Problem.|
|[Labs.Components.IInputComponent](../../reference/office-mix/labs.components.iinputcomponent.md)|Ermöglicht die Interaktion mit einer Eingabekomponente.|
|[Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)|Eine Instanz einer Eingabekomponente.|
