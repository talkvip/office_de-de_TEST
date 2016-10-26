
# Shapes.AddPolyline Method (Excel)

Creates an open polyline or a closed polygon drawing. Returns a  **[Shape](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)** object that represents the new polyline or polygon.


## Syntax

 _Ausdruck_. **AddPolyline**( ** _SafeArrayOfPoints_** )

 _Ausdruck_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Erforderlich|**Variant**|An array of coordinate pairs that specifies the polyline drawing's vertices.|

### Return Value

Shape


## Remarks

To form a closed polygon, assign the same coordinates to the first and last vertices in the polyline drawing.


## Example

This example adds a triangle to  `myDocument`. Because the first and last points have the same coordinates, the polygon is closed and filled. The color of the triangle's interior will be the same as the default shape's fill color.


```
Dim triArray(1 To 4, 1 To 2) As Single 
triArray(1, 1) = 25 
triArray(1, 2) = 100 
triArray(2, 1) = 100 
triArray(2, 2) = 150 
triArray(3, 1) = 150 
triArray(3, 2) = 50 
triArray(4, 1) = 25 ' Last point has same coordinates as first 
triArray(4, 2) = 100 
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddPolyline triArray
```


## Siehe auch


#### Konzepte


[Shapes Object](f9c6548c-d028-1b70-a11c-c4b45ff19177.md)
#### Weitere Ressourcen


[Shapes Object Members](http://msdn.microsoft.com/library/f5d0be42-46cc-2916-8953-401e50a5cef7%28Office.15%29.aspx)