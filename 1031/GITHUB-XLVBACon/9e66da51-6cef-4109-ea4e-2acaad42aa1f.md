
# Range.EntireRow Property (Excel)

Returns a  **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** object that represents the entire row (or rows) that contains the specified range. Read-only.


## Syntax

 _Ausdruck_. **EntireRow**

 _Ausdruck_ A variable that represents a **Range** object.


## Example

This example sets the value of the first cell in the row that contains the active cell. The example must be run from a worksheet.


```
ActiveCell.EntireRow.Cells(1, 1).Value = 5
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example sorts all the rows on a worksheet, including hidden rows.




```
Sub SortAll()
    'Turn off screen updating, and define your variables.
    Application.ScreenUpdating = False
    Dim lngLastRow As Long, lngRow As Long
    Dim rngHidden As Range
    
    'Determine the number of rows in your sheet, and add the header row to the hidden range variable.
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set rngHidden = Rows(1)
    
    'For each row in the list, if the row is hidden add that that row to the hidden range variable.
    For lngRow = 1 To lngLastRow
        If Rows(lngRow).Hidden = True Then
            Set rngHidden = Union(rngHidden, Rows(lngRow))
        End If
    Next lngRow
    
    'Unhide everything in the hidden range variable.
    rngHidden.EntireRow.Hidden = False
    
    'Perform the sort on all the data.
    Range("A1").CurrentRegion.Sort _
        key1:=Range("A2"), _
        order1:=xlAscending, _
        header:=xlYes
        
    'Re-hide the rows that were originally hidden, but unhide the header.
    rngHidden.EntireRow.Hidden = True
    Rows(1).Hidden = False
    
    'Turn screen updating back on.
    Set rngHidden = Nothing
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books veröffentlicht unterhaltsame Bücher für Benutzer von Microsoft Office. Den kompletten Katalog finden Sie unter MrExcel.com.


## Siehe auch
<a name="AboutContributor"> </a>


#### Konzepte


[Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Weitere Ressourcen


[Range Object Members](http://msdn.microsoft.com/library/4336bf81-1e63-7e44-1792-baf366a027a7%28Office.15%29.aspx)