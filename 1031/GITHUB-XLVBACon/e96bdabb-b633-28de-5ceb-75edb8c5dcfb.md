
# Shape.ScaleWidth Method (Excel)

Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.


## Syntax

 _Ausdruck_. **ScaleWidth**( ** _Factor_**, ** _RelativeToOriginalSize_**, ** _Scale_** )

 _Ausdruck_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Factor_|Erforderlich|**Single**|Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Erforderlich|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|**False** to scale it relative to its current size. You can specify **True** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional|**Variant**|One of the constants of  **[MsoScaleFrom](http://msdn.microsoft.com/library/9d1bd699-261a-c360-f680-ff4fac667a31%28Office.15%29.aspx)** which specifies which part of the shape retains its position when the shape is scaled.|

## Remarks




||
|:-----|
|**MsoTriState** can be one of these **MsoTriState** constants.|
|**msoCTrue**. Does not apply to this property.|
|**msoFalse**. To scale it relative to its current size.|
|**msoTriStateMixed**. Does not apply to this property.|
|**msoTriStateToggle**. Does not apply to this property.|
|**msoTrue**. Can only use this argument if the specified shape is a picture or an OLE object.|

## Example

This example scales all pictures and OLE objects on  `myDocument` to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    Select Case s.Type 
    Case msoEmbeddedOLEObject, _ 
            msoLinkedOLEObject, _ 
            msoOLEControlObject, _ 
            msoLinkedPicture, msoPicture 
        s.ScaleHeight 1.75, msoTrue 
        s.ScaleWidth 1.75, ,msoTrue 
    Case Else 
        s.ScaleHeight 1.75, msoFalse 
        s.ScaleWidth 1.75, msoFalse 
    End Select 
Next
```


## Siehe auch


#### Konzepte


[Shape Object](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)
#### Weitere Ressourcen


[Shape Object Members](http://msdn.microsoft.com/library/0fed7136-4228-6c32-507d-3bd36aa56d9a%28Office.15%29.aspx)