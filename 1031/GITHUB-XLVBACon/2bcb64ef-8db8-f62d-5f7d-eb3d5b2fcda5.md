
# PivotTable.VisualTotals Property (Excel)

 **True** (default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/write **Boolean**.


## Syntax

 _Ausdruck_. **VisualTotals**

 _Ausdruck_ A variable that represents a **PivotTable** object.


## Remarks

In non-OLAP PivotTables, if you hide an item, the total is recomputed to reflect only the items that remain visible in the PivotTable. In an OLAP PivotTable, the total is computed on the server and is therefore unaffected by whether any items are hidden in the PivotTable view. However, if the  **VisualTotals** property is set to **False** for an OLAP PivotTable, then the results of the OLAP PivotTable will match those of the non-OLAP PivotTable.

For OLAP PivotTables, a  **VisualTotals** property setting of **True** (default) works the same way as described for non-OLAP PivotTables.

The  **VisualTotals** property returns **True** for all new PivotTables. However, if you open a workbook in the current version of Microsoft Excel and the PivotTable had been created in a previous version of Excel, then the **VisualTotals** property will return **False**.


 **Hinweis**  All previously created PivotTables will have the  **VisualTotals** property set to **False** by default, unless the user changes it, but for all newly created ones the **VisualTotals** property is set to **True**.


## Example

This example determines if the ability to re-total after an item has been hidden from view is available for OLAP PivotTables and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```
Sub CheckVisualTotals() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if visual totals is enabled for OLAP PivotTables. 
 If pvtTable.VisualTotals = True Then 
 MsgBox "Ability enabled to re-total after an item " &amp; _ 
 "has been hidden from view." 
 Else 
 MsgBox "Unable to re-total items not hidden from view." 
 End If 
 
End Sub
```


## Siehe auch


#### Konzepte


[PivotTable Object](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)
#### Weitere Ressourcen


[PivotTable Object Members](http://msdn.microsoft.com/library/8e8d1692-cf32-63c6-a1f6-54ddcc2a4964%28Office.15%29.aspx)