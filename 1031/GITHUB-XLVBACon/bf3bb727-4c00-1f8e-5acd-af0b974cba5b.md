
# PivotFilters.Add2 Method (Excel)

Adds new filters to the  **PivotFilters** collection.


## Syntax

 _Ausdruck_. **Add2**( ** _Type_**, ** _DataField_**, ** _Value1_**, ** _Value2_**, ** _Order_**, ** _Name_**, ** _Description_**, ** _MemberPropertyField_**, ** _WholeDayFilter_** )

 _Ausdruck_ A variable that represents a **PivotFilters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Erforderlich|**XlPivotFilterType**|Requires an  **[XlPivotFilterType](0ae3f0fe-02e3-b0f7-1506-1961c4adcd6c.md)** type of filter.|
| _DataField_|Optional|**Variant**|The field to which the filter is attached.|
| _Value1_|Optional|**Variant**|Filter value 1.|
| _Value2_|Optional|**Variant**|Filter value 2.|
| _Order_|Optional|**Variant**|Order in which the data should be filtered.|
| _Name_|Optional|**Variant**|Name of the filter.|
| _Description_|Optional|**Variant**|A brief description of the filter.|
| _MemberPropertyField_|Optional|**Variant**|Specifies the member property field on which the label filter is based.|
| _WholeDayFilter_|Optional|**Variant**|Specifies a filter based on days.|

### Return Value

PivotFilter


## Example

Following are some examples of how to use the  **Add** function correctly.


```
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlTopCount DataField := MyPivotField2 Value1 := 10 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlCaptionIsNotBetween Value1 := "A" Value2 := "G" 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := 10000 

```

The following example returns a run-time error because the data type of  _Value1_ is invalid.




```
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := ?Allan?
```


## Siehe auch


#### Konzepte


[PivotFilters Object](fc647acb-bd6a-8544-6411-1f5e49807e53.md)
#### Weitere Ressourcen


[PivotFilters Object Members](http://msdn.microsoft.com/library/57f1f375-1b7b-c488-c236-91ed26a68bb6%28Office.15%29.aspx)