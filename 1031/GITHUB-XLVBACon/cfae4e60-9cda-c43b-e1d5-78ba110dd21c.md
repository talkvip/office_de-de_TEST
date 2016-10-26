
# CategoryCollection.Application Property (Excel)

Returns an  **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _Ausdruck_. **Application**

 _Ausdruck_ A variable that represents a[CategoryCollection Object (Excel)](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md) object.


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


## Property value

 **APPLICATION**


## Siehe auch


#### Weitere Ressourcen


[CategoryCollection Object Members](http://msdn.microsoft.com/library/39a6f85c-2219-79df-cbbc-0bcc21a517e8%28Office.15%29.aspx)
[CategoryCollection Object](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md)