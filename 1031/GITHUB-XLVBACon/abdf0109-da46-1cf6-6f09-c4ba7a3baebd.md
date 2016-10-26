
# PivotField.Calculation Property (Excel)

Returns or sets a  **[XlPivotFieldCalculation](94ceaa53-9dfc-149a-6aed-30d8fdb57b5b.md)** value that represents the type of calculation performed by the specified field. This property is valid only for data fields.


## Syntax

 _Ausdruck_. **Calculation**

 _Ausdruck_ A variable that represents a **PivotField** object.


## Example

This example sets the data field in the PivotTable report on Sheet1 to calculate the difference from the base field, sets the base field to the field named "ORDER_DATE," and then sets the base item to the item named "5/16/89."


```
With Worksheets("Sheet1").Range("A3").PivotField 
    .Calculation = xlDifferenceFrom 
    .BaseField = "ORDER_DATE" 
    .BaseItem = "5/16/89" 
End With
```


## Siehe auch


#### Konzepte


[PivotField Object](52784960-e2da-b43a-1e37-2d4dae61c6d8.md)
#### Weitere Ressourcen


[PivotField Object Members](http://msdn.microsoft.com/library/4a6ea12a-072c-a386-c855-7bf5f6eadd46%28Office.15%29.aspx)