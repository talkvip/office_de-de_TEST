
# HPageBreak.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _Ausdruck_. **Application**

 _Ausdruck_ A variable that represents a **HPageBreak** object.


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


[HPageBreak Object](8fc96958-33ab-8251-f627-4769b5eab97f.md)
#### Weitere Ressourcen


[HPageBreak Object Members](http://msdn.microsoft.com/library/32b561ff-a0cf-142b-0a46-c622a42b6125%28Office.15%29.aspx)