
# PivotField.SubtotalName Property (Excel)

Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string "Subtotal". Read/write  **String**.


## Syntax

 _Ausdruck_. **SubtotalName**

 _Ausdruck_ A variable that represents a **PivotField** object.


## Example

This example sets the subtotal label to "Regional Subtotal" (instead of the default string "Subtotal") in the state field in the second PivotTable report on the active worksheet.


```
ActiveSheet.PivotTables("PivotTable2") _ 
 .PivotFields("state").SubtotalName = "Regional Subtotal"
```


## Siehe auch


#### Konzepte


[PivotField Object](52784960-e2da-b43a-1e37-2d4dae61c6d8.md)
#### Weitere Ressourcen


[PivotField Object Members](http://msdn.microsoft.com/library/4a6ea12a-072c-a386-c855-7bf5f6eadd46%28Office.15%29.aspx)