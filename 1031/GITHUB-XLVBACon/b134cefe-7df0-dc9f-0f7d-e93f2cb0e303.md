
# CubeField.Orientation Property (Excel)

Returns or sets a  **[XlPivotFieldOrientation](8dd82d0c-370a-464f-e666-5bc8cbcdacb4.md)** value that represents the location of the field in the specified PivotTable report.


## Syntax

 _Ausdruck_. **Orientation**

 _Ausdruck_ A variable that represents a **CubeField** object.


## Remarks

For OLAP data sources, setting this property for one field in a hierarchy sets the orientation for the other fields in the same hierarchy. Dimension fields can only be oriented in the row, column, and page field areas of the PivotTable report. Measure fields can only be oriented in the data area. Setting a hierarchy or data field to  **xlHidden** removes the hierarchy or field from the PivotTable report.


## Siehe auch


#### Konzepte


[CubeField Object](6db16910-6c27-651a-c388-e54e27fe4519.md)
#### Weitere Ressourcen


[CubeField Object Members](http://msdn.microsoft.com/library/2f3cbe65-45ff-abe0-3e48-29c0d490f600%28Office.15%29.aspx)