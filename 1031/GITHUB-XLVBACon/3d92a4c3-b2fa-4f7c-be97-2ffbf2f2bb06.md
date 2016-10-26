
# Füllen eines Wert nach unten in leere Zellen in einer Spalte

Das folgende Beispiel betrachtet Spalte A. Wenn sich dort eine leere Zelle befindet, wird der Wert der leeren Zelle auf den gleichen Wert der Zelle darüber festgelegt. Dies wird für die gesamte Spalte durchgeführt, sodass in jede leere Zelle ein Wert eingefügt wird.

 **Beispielcode bereitgestellt von:** Tom Urtis,[Atlas Programming Management](http://www.atlaspm.com/)



```
Sub FillCellsFromAbove()
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next
    ' Look in column A
    With Columns(1)
        ' For blank cells, set them to equal the cell above
        .SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        'Convert the formula to a value
        .Value = .Value
    End With
    Err.Clear
    Application.ScreenUpdating = True
End Sub
```


## Über den Autor
<a name="AboutContributor"> </a>

MVP Tom Urtis ist Gründer von Atlas Programming Management, einem Full-Service-Unternehmen für Microsoft Office- und Excel-Geschäftslösungen mit Sitz im kalifornischen Silicon Valley. Urtis verfügt über mehr als 25 Jahre Erfahrung in der Unternehmensführung und der Entwicklung von Microsoft Office-Anwendungen und ist Co-Author von "Holy Macro! It's 2,500 Excel VBA Examples".

