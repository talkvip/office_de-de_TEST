
# FormatCondition.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _Ausdruck_. **Application**

 _Ausdruck_ A variable that represents a **FormatCondition** object.


## Example

This example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Siehe auch


#### Konzepte


[FormatCondition Object](38a2bca9-9b28-3ef2-8c7a-4d35a27229ec.md)
#### Weitere Ressourcen


[FormatCondition Object Members](http://msdn.microsoft.com/library/8f4bebce-0bf4-03de-62f0-4454ea699c5f%28Office.15%29.aspx)