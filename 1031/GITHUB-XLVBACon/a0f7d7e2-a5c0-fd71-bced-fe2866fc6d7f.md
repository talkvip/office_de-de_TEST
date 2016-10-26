
# Shape.FormControlType Property (Excel)

Returns the Microsoft Excel control type. Read-only  **[XlFormControl](fad54f9d-4ef2-b2ac-d187-131e5bdd18e1.md)**.


## Syntax

 _Ausdruck_. **FormControlType**

 _Ausdruck_ A variable that represents a **Shape** object.


## Remarks

You cannot use this property with ActiveX controls (the  **[Type](93939e9f-2630-4db2-6b66-6705720877f6.md)** property for the **[Shape](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)** object must return **msoFormControl** ).


## Example

This example clears all the Microsoft Excel check boxes on worksheet one.


```
For Each s In Worksheets(1).Shapes 
 If s.Type = msoFormControl Then 
 If s.FormControlType = xlCheckBox Then _ 
 s.ControlFormat.Value = False 
 End If 
Next
```


## Siehe auch


#### Konzepte


[Shape Object](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)
#### Weitere Ressourcen


[Shape Object Members](http://msdn.microsoft.com/library/0fed7136-4228-6c32-507d-3bd36aa56d9a%28Office.15%29.aspx)