
# ShapeRange.IncrementTop Method (Excel)

Moves the specified shape vertically by the specified number of points.


## Syntax

 _Ausdruck_. **IncrementTop**( ** _Increment_** )

 _Ausdruck_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Erforderlich|**Single**|Specifies how far the shape object is to be moved vertically, in points. A positive value moves the shape down; a negative value moves it up.|

## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## Siehe auch


#### Konzepte


[ShapeRange Object](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)
#### Weitere Ressourcen


[ShapeRange Object Members](http://msdn.microsoft.com/library/1d1950c5-32ac-dfc0-8c19-07159a29a2a0%28Office.15%29.aspx)