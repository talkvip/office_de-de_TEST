
# Selecionando e ativando células

No Microsoft Excel, geralmente você seleciona uma célula ou células e então executa uma ação, como formatar as células ou inserir valores nelas. No Visual Basic, normalmente não é necessário selecionar células antes de modificá-las.

Por exemplo, para inserir uma fórmula na célula D6 usando o Visual Basic, não é necessário selecionar o intervalo D6. Basta retornar o objeto  **Range** para essa célula e definir a propriedade **Formula** para a fórmula desejada, como mostrado no exemplo a seguir.



```VB.net
Sub EnterFormula() 
    Worksheets("Sheet1").Range("D6").Formula = "=SUM(D2:D5)" 
End Sub
```

Para mais informações e exemplos de como usar outros métodos para controlar células sem selecioná-las, consulte [Como fazer referência a células e a intervalos](a16caa8d-21c9-ff33-347b-ce671248a92d.md).

## Usando o método Select e a propriedade Selection

O método  **Select** ativa planilhas e objetos em planilhas; a propriedade **Selection** retorna um objeto que representa a seleção atual na planilha ativa na pasta de trabalho ativa. Antes de poder usar a propriedade **Selection** com êxito, você deverá ativar uma pasta de trabalho, ativar ou selecionar uma planilha e selecionar um intervalo (ou outro objeto) usando o método **Select**.

Com frequência, o gravador de macros criará uma macro que usa o método  **Select** e a propriedade **Selection**. O **Sub** procedimento foi criado usando o gravador de macros e mostra como **Select** e **Selection** funcionam juntos.




```VB.net
Sub Macro1() 
    Sheets("Sheet1").Select 
    Range("A1").Select 
    ActiveCell.FormulaR1C1 = "Name" 
    Range("B1").Select 
    ActiveCell.FormulaR1C1 = "Address" 
    Range("A1:B1").Select 
    Selection.Font.Bold = True 
End Sub
```

O exemplo a seguir executa a mesma tarefa sem ativar nem selecionar a planilha ou as células.




```VB.net
Sub Labels() 
    With Worksheets("Sheet1") 
        .Range("A1") = "Name" 
        .Range("B1") = "Address" 
        .Range("A1:B1").Font.Bold = True 
    End With 
End Sub
```


## Selecionar células na planilha ativa

Se você usar o método  **Select** para selecionar células, lembre-se de que **Select** funciona apenas na planilha ativa. Se você executar o procedimento **Sub** do módulo, o método **Select** falhará, a menos que o procedimento ative a planilha antes de usar o método **Select** em um intervalo de células. Por exemplo, o procedimento a seguir copia uma linha da Planilha1 para a Planilha2 na pasta de trabalho ativa.


```VB.net
Sub CopyRow() 
    Worksheets("Sheet1").Rows(1).Copy 
    Worksheets("Sheet2").Select 
    Worksheets("Sheet2").Rows(1).Select 
    Worksheets("Sheet2").Paste 
End Sub
```


## Ativar uma célula dentro de uma seleção

Você pode usar o método  **Activate** para ativar uma célula dentro de uma seleção. Pode haver apenas uma célula ativa, mesmo quando um intervalo de células é selecionado. O procedimento a seguir seleciona um intervalo e ativa uma célula dentro do intervalo sem alterar a seleção.


```VB.net
Sub MakeActive() 
    Worksheets("Sheet1").Activate 
    Range("A1:D4").Select 
    Range("B2").Activate 
End Sub
```

