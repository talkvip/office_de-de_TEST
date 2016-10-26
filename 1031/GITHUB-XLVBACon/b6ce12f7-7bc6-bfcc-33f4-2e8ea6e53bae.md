
# Workbook.PivotTableOpenConnection Event (Excel)

Occurs after a PivotTable report opens the connection to its data source.


## Syntax

 _Ausdruck_. **PivotTableOpenConnection**( ** _Target_** )

 _Ausdruck_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Erforderlich|**[PivotTable](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)**|The selected PivotTable report.|
| _Target_|Erforderlich|PIVOTTABLE||
|Name|Required/Optional|Data type|Description|

### Return Value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been opened. This example assumes you have declared an object of type  **Workbook** with events in a class module.


```
Private Sub ConnectionApp_PivotTableOpenConnection(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been opened." 
 
End Sub
```


## Siehe auch


#### Konzepte


[Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Weitere Ressourcen


[Workbook Object Members](http://msdn.microsoft.com/library/dce102a3-25de-3ff4-2ce5-bc56e08baca7%28Office.15%29.aspx)