
# Shapes.AddLabel Method (Excel)

Creates a label. Returns a  **[Shape](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)** object that represents the new label.


## Syntax

 _Ausdruck_. **AddLabel**( ** _Orientation_**, ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _Ausdruck_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Erforderlich|**[MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The text orientation within the label.|
| _Left_|Erforderlich|**Single**|The position (in points) of the upper-left corner of the label relative to the upper-left corner of the document.|
| _Top_|Erforderlich|**Single**|The position (in points) of the upper-left corner of the label relative to the top corner of the document.|
| _Width_|Erforderlich|**Single**|The width of the label, in points.|
| _Height_|Erforderlich|**Single**|The height of the label, in points.|

### Return Value

Shape


## Example

This example adds a vertical label that contains the text "Test Label" to  `myDocument`.


```
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddLabel(msoTextOrientationVertical, _ 
    100, 100, 60, 150) _ 
    .TextFrame.Characters.Text = "Test Label"
```


## Siehe auch


#### Konzepte


[Shapes Object](f9c6548c-d028-1b70-a11c-c4b45ff19177.md)
#### Weitere Ressourcen


[Shapes Object Members](http://msdn.microsoft.com/library/f5d0be42-46cc-2916-8953-401e50a5cef7%28Office.15%29.aspx)