
# Labs.Components.IInputComponent

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Enables interacting with an input component.

```
interface IInputComponent extends Labs.Core.IComponent
```


## Properties


|||
|:-----|:-----|
| `maxScore: number`|The maximum allowable score for the input component.|
| `timeLimit: number`|Time limit for the input problem.|
| `hasAnswer: boolean`|**True** if the component has an answer.|
| `answer: any`|The answer to the component problem, if any.|
| `secure: boolean`|**True** if the input component is secure.|
