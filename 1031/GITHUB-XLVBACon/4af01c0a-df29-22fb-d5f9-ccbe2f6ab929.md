
# ODBCConnection.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

 _Ausdruck_. **Creator**

 _Ausdruck_ A variable that represents an **ODBCConnection** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## Siehe auch


#### Konzepte


[ODBCConnection Object](b880ebec-15a4-5a3d-ef02-db73106db9c9.md)
#### Weitere Ressourcen


[ODBCConnection Object Members](http://msdn.microsoft.com/library/d13b91f3-a89f-7dd7-7a98-f1d952f3b047%28Office.15%29.aspx)