
# Outline.ShowLevels Method (Excel)

Displays the specified number of row and/or column levels of an outline.


## Syntax

 _Ausdruck_. **ShowLevels**( ** _RowLevels_**, ** _ColumnLevels_** )

 _Ausdruck_ A variable that represents an **Outline** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowLevels_|Optional|**Variant**|Specifies the number of row levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on rows.|
| _ColumnLevels_|Optional|**Variant**|Specifies the number of column levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on columns.|

### Return Value

Variant


## Remarks

You must specify at least one argument.


## Example

This example displays row levels one through three and column level one of the outline on Sheet1.


```
Worksheets("Sheet1").Outline _ 
 .ShowLevels rowLevels:=3, columnLevels:=1
```


## Siehe auch


#### Konzepte


[Outline Object](f5d50a8a-0dd9-638a-4374-5c648386a598.md)
#### Weitere Ressourcen


[Outline Object Members](http://msdn.microsoft.com/library/bf8e2103-d023-fc1f-90f2-960dff36e548%28Office.15%29.aspx)