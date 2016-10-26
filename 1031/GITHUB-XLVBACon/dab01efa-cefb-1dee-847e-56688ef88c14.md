
# OLEDBError.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

 _Ausdruck_. **Creator**

 _Ausdruck_ A variable that represents an **OLEDBError** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## Siehe auch


#### Konzepte


[OLEDBError Object](6bcbf721-f2c8-f784-361b-e1a298bb2ecb.md)
#### Weitere Ressourcen


[OLEDBError Object Members](http://msdn.microsoft.com/library/52181252-dd6f-b267-fa21-4ad8175b7346%28Office.15%29.aspx)