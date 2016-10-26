
# AutoCorrect.AutoExpandListRange Property (Excel)

A  **Boolean** value indicating whether automatic expansion is enabled for lists. When you type in a cell of an empty row or column next to a list, the list will expand to include that row or column if automatic expansion is enabled. Read/write **Boolean**.


## Syntax

 _Ausdruck_. **AutoExpandListRange**

 _Ausdruck_ A variable that represents an **AutoCorrect** object.


## Example

The following example enables automatic expansion of lists when typing in adjacent rows or columns.


```
Sub SetAutoExpand 
 
 Application.AutoCorrect.AutoExpandListRange = TRUE 
 
End Sub
```


## Siehe auch


#### Konzepte


[AutoCorrect Object](2594722a-2ff9-7175-4d35-0da0ad413b0d.md)
#### Weitere Ressourcen


[AutoCorrect Object Members](http://msdn.microsoft.com/library/ee525804-da41-f613-3e2a-6f6b115dcdd6%28Office.15%29.aspx)