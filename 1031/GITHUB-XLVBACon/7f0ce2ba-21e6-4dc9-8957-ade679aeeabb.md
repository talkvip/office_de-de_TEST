
# HPageBreak.Location Property (Excel)

Returns or sets the cell (a  **Range** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/write **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)**.


## Syntax

 _Ausdruck_. **Location**

 _Ausdruck_ A variable that represents a **HPageBreak** object.


## Example

 This example sets the horizontal page-break location. Note that you must be in Page Break Preview mode in order to set it.


```
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```

 **Note:** **HPageBreak.Location** can only be used to set the horizontal page-break location. In order to change the location of a **HPageBreak**, you must use[HPageBreak.Dragoff](80065224-c53d-3f45-8d94-c644502dac22.md).


## Siehe auch


#### Konzepte


[HPageBreak Object](8fc96958-33ab-8251-f627-4769b5eab97f.md)
#### Weitere Ressourcen


[HPageBreak Object Members](http://msdn.microsoft.com/library/32b561ff-a0cf-142b-0a46-c622a42b6125%28Office.15%29.aspx)