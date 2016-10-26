
# Workbook.Save Method (Excel)

Saves changes to the specified workbook.


## Syntax

 _Ausdruck_. **Save**

 _Ausdruck_ A variable that represents a **Workbook** object.


## Remarks

To open a workbook file, use the  **[Open](1d1c3fca-ae1a-0a91-65a2-6f3f0fb308a0.md)** method.

To mark a workbook as saved without writing it to a disk, set its  **[Saved](37eb8e08-2bfa-8065-2520-a71e291ab50c.md)** property to **True**.

The first time you save a workbook, use the  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** method to specify a name for the file.


## Example

This example saves the active workbook.


```
ActiveWorkbook.Save
```

This example saves all open workbooks and then closes Microsoft Excel.




```
For Each w In Application.Workbooks 
    w.Save 
Next w 
Application.Quit
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example uses the  **BeforeSave** event to verify that certain cells contain data before the workbook can be saved. The workbook cannot be saved until there is data in each of the following cells: D5, D7, D9, D11, D13, and D15.




```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
   'If the six specified cells do not contain data, then display a message box with an error
   'and cancel the attempt to save.
   If WorksheetFunction.CountA(Worksheets("Sheet1").Range("D5,D7,D9,D11,D13, D15")) < 6 Then
      MsgBox "Workbook will not be saved unless" &amp; vbCrLf &amp; _
      "All required fields have been filled in!"
      Cancel = True
   End If
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books veröffentlicht unterhaltsame Bücher für Benutzer von Microsoft Office. Den kompletten Katalog finden Sie unter MrExcel.com.


## Siehe auch
<a name="AboutContributor"> </a>


#### Konzepte


[Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Weitere Ressourcen


[Workbook Object Members](http://msdn.microsoft.com/library/dce102a3-25de-3ff4-2ce5-bc56e08baca7%28Office.15%29.aspx)