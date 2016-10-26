
# Range.Consolidate Method (Excel)

Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.  **Variant**.


## Syntax

 _Ausdruck_. **Consolidate**( ** _Sources_**, ** _Function_**, ** _TopRow_**, ** _LeftColumn_**, ** _CreateLinks_** )

 _Ausdruck_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sources_|Optional|**Variant**|The sources of the consolidation as an array of text reference strings in R1C1-style notation. The references must include the full path of sheets to be consolidated.|
| _Function_|Optional|**Variant**|One of the constants of  **[XlConsolidationFunction](a3d0e4c0-8463-340c-a258-219d49f715d7.md)** which specifies the type of consolidation.|
| _TopRow_|Optional|**Variant**|**True** to consolidate data based on column titles in the top row of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
| _LeftColumn_|Optional|**Variant**|**True** to consolidate data based on row titles in the left column of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
| _CreateLinks_|Optional|**Variant**|**True** to have the consolidation use worksheet links. **False** to have the consolidation copy the data. The default value is **False**.|

### Return Value

Variant


## Example

This example consolidates data from Sheet2 and Sheet3 onto Sheet1, using the SUM function.


```
Worksheets("Sheet1").Range("A1").Consolidate _ 
 Sources:=Array("Sheet2!R1C1:R37C6", "Sheet3!R1C1:R37C6"), _ 
 Function:=xlSum
```


## Siehe auch


#### Konzepte


[Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Weitere Ressourcen


[Range Object Members](http://msdn.microsoft.com/library/4336bf81-1e63-7e44-1792-baf366a027a7%28Office.15%29.aspx)