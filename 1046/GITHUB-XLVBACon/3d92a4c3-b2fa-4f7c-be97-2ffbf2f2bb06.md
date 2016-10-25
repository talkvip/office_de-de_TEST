
# Preencher as células em branco com um valor em uma coluna

O exemplo a seguir examina a coluna A e, se houver uma célula em branco, define o valor da célula em branco como o valor da célula acima dela. Isso continua por toda a coluna até a parte inferior, preenchendo com um valor cada célula em branco.

 **Código de exemplo fornecido por:** Tom Urtis,[Atlas Programming Management](http://www.atlaspm.com/)



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


## Sobre o colaborador
<a name="AboutContributor"> </a>

MVP Tom Urtis é o fundador da Atlas Programming Management, uma completa do serviço Microsoft Office e o Excel business solutions empresa no Silicon Valley. João tem mais de 25 anos de experiência em gerenciamento de negócios e desenvolvimento Microsoft Office aplicativos e é co-autor de "Holy Macro! It's 2,500 Excel VBA Examples."

