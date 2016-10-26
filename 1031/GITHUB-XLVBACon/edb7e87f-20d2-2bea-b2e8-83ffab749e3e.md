
# Style.IncludePatterns Property (Excel)

 **True** if the style includes the **Color**, **ColorIndex**, **InvertIfNegative**, **Pattern**, **PatternColor**, and **PatternColorIndex** interior properties. Read/write **Boolean**.


## Syntax

 _Ausdruck_. **IncludePatterns**

 _Ausdruck_ A variable that represents a **Style** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include pattern format.


```
Worksheets("Sheet1").Range("A1").Style.IncludePatterns = True
```


## Siehe auch


#### Konzepte


[Style Object](3c1e9184-0075-5f46-9a1a-0b61d874d1f8.md)
#### Weitere Ressourcen


[Style Object Members](http://msdn.microsoft.com/library/78f477c9-4033-e7c5-fc3d-7ba025392d31%28Office.15%29.aspx)