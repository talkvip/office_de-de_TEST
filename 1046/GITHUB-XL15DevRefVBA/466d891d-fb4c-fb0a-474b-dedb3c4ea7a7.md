
# Método Workbook.Save (Excel)

Salva as alterações na pasta de trabalho especificada.


## Sintaxe

 _expressão_. **Save**

 _expressão_ Uma variável que representa um objeto **Workbook**.


## Comentários

Para abrir um arquivo de pasta de trabalho, use o método  **[Open](1d1c3fca-ae1a-0a91-65a2-6f3f0fb308a0.md)**.

Para marcar uma pasta de trabalho como salva sem gravá-la em um disco, defina sua propriedade  **[Saved](37eb8e08-2bfa-8065-2520-a71e291ab50c.md)** como **True**.

Na primeira vez em que você salvar uma pasta de trabalho, use o método  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** para especificar um nome para o arquivo.


## Exemplo

Este exemplo salva a pasta de trabalho ativa.


```
ActiveWorkbook.Save
```

Este exemplo salva todas as pastas de trabalho abertas e então fecha o Microsoft Excel.




```
For Each w In Application.Workbooks 
    w.Save 
Next w 
Application.Quit
```

 **Exemplo de código fornecido por:** Holy Macro! Books,[Santa Macro! São 2.500 exemplos do VBA do Excel](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

Este exemplo usa o evento  **BeforeSave** para verificar se determinadas células contêm dados antes que a pasta de trabalho possa ser salva. A pasta de trabalho não poderá ser salva até haver dados em cada uma destas células: D5, D7, D9, D11, D13 e D15.




```
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
   'If the six specified cells do not contain data, then display a message box with an error
   'and cancel the attempt to save.
   If WorksheetFunction.CountA(Worksheets("Sheet1").Range("D5,D7,D9,D11,D13, D15")) < 6 Then
      MsgBox "Workbook will not be saved unless" &amp; vbCrLf &amp; _
      "All required fields have been filled in!"
      Cancel = True
   End If
End Sub
```


## Sobre o colaborador
<a name="AboutContributor"> </a>

Holy Macro! Livros publica livros de entretenimento para as pessoas que usam o Microsoft Office. Veja o catálogo completo em MrExcel.com.


## Ver também
<a name="AboutContributor"> </a>


#### Conceitos


[Objeto Workbook](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Outros recursos


[Membros do objeto Workbook](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)