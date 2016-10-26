
# Application.WorkbookAfterSave Event (Excel)

Occurs after the workbook is saved.


## Syntax

 _Ausdruck_. **WorkbookAfterSave**( ** _Wb_**, ** _Success_** )

 _Ausdruck_ A variable that represents an **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Erforderlich|**[Workbook](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)**|The workbook being saved.|
| _Success_|Erforderlich|**Boolean**|Returns  **True** if the save operation was successful; otherwise **False**.|

### Return Value

Nothing


## Example

The following code example displays a message box if the workbook was successfully saved. This code must be placed in a class module and an instance of that class must be correctly initialized. For more information about how to use event procedures with the  **Application** object, see[Using Events with the Application Object](0063feba-47fd-29be-d2d5-8fcf47e70cbc.md).


```
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " &amp; Wb.Name &amp; " workbook was successfully saved.") 
End If 
End Sub
```


## Siehe auch


#### Konzepte


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Weitere Ressourcen


[Application Object Members](http://msdn.microsoft.com/library/4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e%28Office.15%29.aspx)