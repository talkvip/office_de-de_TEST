
# Window.DisplayZeros Property (Excel)

 **True** if zero values are displayed. Read/write **Boolean**.


## Syntax

 _Ausdruck_. **DisplayZeros**

 _Ausdruck_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.


## Example

This example sets the active window in Book1.xls to display zero values.


```
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayZeros = True 

```


## Siehe auch


#### Konzepte


[Window Object](8591b1ad-76f8-14e2-9120-406b65093f5a.md)
#### Weitere Ressourcen


[Window Object Members](http://msdn.microsoft.com/library/f11db427-24a4-041c-2fd5-03ce73ae6c16%28Office.15%29.aspx)