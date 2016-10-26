
# ShapeRange.TextEffect Property (Excel)

Returns a  **[TextEffectFormat](7fe03721-6a45-569e-add4-fc8849c99535.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _Ausdruck_. **TextEffect**

 _Ausdruck_ A variable that represents a **ShapeRange** object.


## Example

This example sets the font style to bold for shape three on  `myDocument` if the shape is WordArt.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = True 
 End If 
End With
```


## Siehe auch


#### Konzepte


[ShapeRange Object](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)
#### Weitere Ressourcen


[ShapeRange Object Members](http://msdn.microsoft.com/library/1d1950c5-32ac-dfc0-8c19-07159a29a2a0%28Office.15%29.aspx)