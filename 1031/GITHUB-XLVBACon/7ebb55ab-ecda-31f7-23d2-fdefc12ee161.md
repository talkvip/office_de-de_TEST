
# PivotTable.ShowPages Method (Excel)

Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.


## Syntax

 _Ausdruck_. **ShowPages**( ** _PageField_** )

 _Ausdruck_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageField_|Optional|**Variant**|A string that names a single page field in the report.|

### Return Value

Variant


## Remarks

This method isn?t available for OLAP data sources.


## Example

This example creates a new PivotTable report for each item in the page field, which is the field named ?Country.?


```
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.ShowPages "Country"
```


## Siehe auch


#### Konzepte


[PivotTable Object](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)
#### Weitere Ressourcen


[PivotTable Object Members](http://msdn.microsoft.com/library/8e8d1692-cf32-63c6-a1f6-54ddcc2a4964%28Office.15%29.aspx)