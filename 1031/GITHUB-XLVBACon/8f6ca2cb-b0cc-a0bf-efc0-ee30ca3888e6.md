
# SparklineGroup.ModifyLocation Method (Excel)

Sets the associated  **[Range](8bc4841b-72f7-34b5-a299-3357bf8f457b.md)** object to modify the location of the sparkline group.


## Syntax

 _Ausdruck_. **ModifyLocation**( ** _Location_** )

 _Ausdruck_ A variable that represents a **[SparklineGroup](cc694d97-a3d3-3473-2e37-0ede67b97680.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Location_|Erforderlich|**Range**|The  **Range** that represents the location of the sparkline group.|

### Return Value

Nothing


## Example

This example selects a sparkline group in the location A1:A4 and changes the location to equal A10:A14.


```
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifyLocation Range("$A$10:$A$14")
```


## Siehe auch


#### Konzepte


[SparklineGroup Object](cc694d97-a3d3-3473-2e37-0ede67b97680.md)
#### Weitere Ressourcen


[SparklineGroup Object Members](http://msdn.microsoft.com/library/dad308ee-d69b-748d-d0c8-ad63c643808f%28Office.15%29.aspx)