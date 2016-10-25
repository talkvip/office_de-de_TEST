
# Método Range.PasteSpecial (Excel)

Cola um  **[Intervalo](b8207778-0dcc-4570-1234-f130532cc8cd.md)** copiado no intervalo especificado.


## Sintaxe

 _expressão_. **PasteSpecial**( ** _Paste_**, ** _Operation_**, ** _SkipBlanks_**, ** _Transpose_** )

 _expressão_ Uma variável que representa um objeto **Range**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Paste_|Opcional|**[XlPasteType](a60202d9-b380-ed88-b7d8-66bf34e032a5.md)**|. A parte do intervalo a ser colada.|
| _Operation_|Opcional|**[XlPasteSpecialOperation](b1e01a39-61b8-a3a9-2552-58d79b10afe3.md)**|. A operação de colagem.|
| _SkipBlanks_|Opcional|**Variant**|**True** para ter células em branco no intervalo na Área de Transferência que não deve ser colada no intervalo de destino. O valor padrão é **False**.|
| _Transpose_|Opcional|**Variant**|**True** para transpor linhas e colunas quando o intervalo é colado.O valor padrão é **False**.|

### Valor de retorno

Variante


## Exemplo

Este exemplo substitui os dados nas células D1:D5 na Planilha1 com a soma do conteúdo existente e das células C1:C5 na Planilha1.


```
With Worksheets("Sheet1") 
 .Range("C1:C5").Copy 
 .Range("D1:D5").PasteSpecial _ 
 Operation:=xlPasteSpecialOperationAdd 
End With
```


## Ver também


#### Conceitos


[Objeto Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Outros recursos


[Membros do objeto Range](4336bf81-1e63-7e44-1792-baf366a027a7.md)