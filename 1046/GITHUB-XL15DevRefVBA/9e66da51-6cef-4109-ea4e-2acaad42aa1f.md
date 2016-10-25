
# Propriedade Range.EntireRow (Excel)

Retorna um objeto  **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** que representa a linha (ou linhas) inteira que contém o intervalo especificado. Somente leitura.


## Sintaxe

 _expressão_. **EntireRow**

 _expressão_ Uma variável que representa um objeto **Range**.


## Exemplo

Este exemplo define o valor da primeira célula na linha que contém a célula ativa. O exemplo deve ser executado de uma planilha.


```
ActiveCell.EntireRow.Cells(1, 1).Value = 5
```

Exemplos

Este exemplo classifica todas as linhas em uma planilha, incluindo as linhas ocultas.




```
Sub SortAll()
    'Turn off screen updating, and define your variables.
    Application.ScreenUpdating = False
    Dim lngLastRow As Long, lngRow As Long
    Dim rngHidden As Range
    
    'Determine the number of rows in your sheet, and add the header row to the hidden range variable.
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set rngHidden = Rows(1)
    
    'For each row in the list, if the row is hidden add that that row to the hidden range variable.
    For lngRow = 1 To lngLastRow
        If Rows(lngRow).Hidden = True Then
            Set rngHidden = Union(rngHidden, Rows(lngRow))
        End If
    Next lngRow
    
    'Unhide everything in the hidden range variable.
    rngHidden.EntireRow.Hidden = False
    
    'Perform the sort on all the data.
    Range("A1").CurrentRegion.Sort _
        key1:=Range("A2"), _
        order1:=xlAscending, _
        header:=xlYes
        
    'Re-hide the rows that were originally hidden, but unhide the header.
    rngHidden.EntireRow.Hidden = True
    Rows(1).Hidden = False
    
    'Turn screen updating back on.
    Set rngHidden = Nothing
    Application.ScreenUpdating = True
End Sub
```


## Sobre o colaborador
<a name="AboutContributor"> </a>

Holy Macro! Livros publica livros de entretenimento para as pessoas que usam o Microsoft Office. Veja o catálogo completo em MrExcel.com.


## Ver também
<a name="AboutContributor"> </a>


#### Conceitos


[Objeto Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Outros recursos


[Membros do objeto Range](4336bf81-1e63-7e44-1792-baf366a027a7.md)