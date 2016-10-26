
# PivotCache.MissingItemsLimit Property (Excel)

Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write  **[XlPivotTableMissingItems](3450ac87-7a30-f2dd-efc8-fcd336b26319.md)**.


## Syntax

 _Ausdruck_. **MissingItemsLimit**

 _Ausdruck_ A variable that represents a **PivotCache** object.


## Remarks


||
|:-----|
|**XlPivotTableMissingItems** can be one of these **XlPivotTableMissingItems** constants.|
|**xlMissingItemsDefault** The default number of unique items per PivotField allowed.|
|**xlMissingItemsMax** The maximum number of unique items per PivotField allowed (32,500).|
|**xlMissingItemsNone** No unique items per PivotField allowed (zero).|
This property can be set to a value between 0 and 32500. If an integer less than zero is specified, this is equivalent to specifying  **xlMissingItemsDefault**. Integers greater than 32,500 can be specified but will have the same effect as specifying **xlMissingItemsMax**.

The  **MissingItemsLimit** property only works for non-OLAP PivotTables; otherwise, a run-time error can occur.


## Example

This example determines the maximum quantity of unique items per field and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```
Sub CheckMissingItemsList() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the maximum number of unique items allowed per PivotField and notify the user. 
 Select Case pvtCache.MissingItemsLimit 
 Case xlMissingItemsDefault 
 MsgBox "The default value of unique items per PivotField is allowed." 
 Case xlMissingItemsMax 
 MsgBox "The maximum value of unique items per PivotField is allowed." 
 Case xlMissingItemsNone 
 MsgBox "No unique items per PivotField are allowed." 
 End Select 
 
End Sub
```


## Siehe auch


#### Konzepte


[PivotCache Object](c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439.md)
#### Weitere Ressourcen


[PivotCache Object Members](http://msdn.microsoft.com/library/113f1109-e1c9-2c6e-0581-9fba82f278dc%28Office.15%29.aspx)