
# FormatConditions.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _Ausdruck_. **Item**( ** _Index_** )

 _Ausdruck_ A variable that represents a **FormatConditions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Erforderlich|**Variant**|The name or index number for the object.|

### Return Value

An Object value that represents an object contained by the collection.


## Example

This example sets format properties for an existing conditional format for cells E1:E10.


```
With Worksheets(1).Range("e1:e10").FormatConditions.Item(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
End With
```


## Siehe auch


#### Konzepte


[FormatConditions Object](2486d4b4-605c-76d8-132a-694c0c600a81.md)
#### Weitere Ressourcen


[FormatConditions Object Members](http://msdn.microsoft.com/library/0e5a3774-fe65-597f-9b97-3bba637b55cc%28Office.15%29.aspx)