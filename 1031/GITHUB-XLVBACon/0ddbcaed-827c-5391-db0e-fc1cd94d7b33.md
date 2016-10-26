
# ShapeRange.SetShapesDefaultProperties Method (Excel)

Makes the formatting of the specified shape the default formatting for the shape.


## Syntax

 _Ausdruck_. **SetShapesDefaultProperties**

 _Ausdruck_ A variable that represents a **ShapeRange** object.


## Example

This example adds a rectangle to  `myDocument`, formats the rectangle's fill, sets the rectangle's formatting as the default shape formatting, and then adds another smaller rectangle to the document. The second rectangle has the same fill as the first one.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 With .AddShape(msoShapeRectangle, 5, 5, 80, 60) 
 With .Fill 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(0, 204, 255) 
 .Patterned msoPatternHorizontalBrick 
 End With 
 ' Set formatting as default formatting 
 .SetShapesDefaultProperties 
 End With 
 ' Create new shape with default formatting 
 .AddShape msoShapeRectangle, 90, 90, 40, 30 
End With
```


## Siehe auch


#### Konzepte


[ShapeRange Object](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)
#### Weitere Ressourcen


[ShapeRange Object Members](http://msdn.microsoft.com/library/1d1950c5-32ac-dfc0-8c19-07159a29a2a0%28Office.15%29.aspx)