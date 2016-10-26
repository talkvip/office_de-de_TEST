
# Shape.AutoShapeType Property (Excel)

Returns or sets the shape type for the specified  **[Shape](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)** or **[ShapeRange](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)**.


## Syntax

 _Ausdruck_. **AutoShapeType**

 _Ausdruck_ A variable that represents a **Shape** object.


## Remarks

When you change the type of a shape, the shape retains its size, color, and other attributes.

Use the  **[Type](55be2527-9fd5-6930-ff64-e3355a36e9e9.md)** property of the **[ConnectorFormat](56c97d73-bde2-52ae-2bc3-724d21fdd515.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in  `myDocument`.


```
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


## Siehe auch


#### Konzepte


[Shape Object](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)
#### Weitere Ressourcen


[Shape Object Members](http://msdn.microsoft.com/library/0fed7136-4228-6c32-507d-3bd36aa56d9a%28Office.15%29.aspx)