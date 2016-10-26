
# Application.DDEAppReturnCode Property (Excel)

Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only  **Long**.


## Syntax

 _Ausdruck_. **DDEAppReturnCode**

 _Ausdruck_ A variable that represents an **Application** object.


## Example

This example sets the variable  `appErrorCode` to the DDE return code.


```
appErrorCode = Application.DDEAppReturnCode
```


## Siehe auch


#### Konzepte


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Weitere Ressourcen


[Application Object Members](http://msdn.microsoft.com/library/4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e%28Office.15%29.aspx)