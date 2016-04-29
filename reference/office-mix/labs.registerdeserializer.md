
# Labs.registerDeserializer

 _**Gilt für:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Deserializes a specified JSON object into an object. Should be used by component authors only.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Parameter


|||
|:-----|:-----|
|json|The [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) to deserialize.|

## Return value

Returns an [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) instance.

