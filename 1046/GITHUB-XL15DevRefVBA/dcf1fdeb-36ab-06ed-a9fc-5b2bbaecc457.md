
# Método Window.SmallScroll (Excel)

Rola o conteúdo da janela por linhas ou colunas.


## Sintaxe

 _expressão_. **SmallScroll**( ** _Down_**, ** _Up_**, ** _ToRight_**, ** _ToLeft_** )

 _expressão_ Uma variável que representa um objeto **Window**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Down_|Opcional|**Variant**|O número de linhas pelas quais rolar o conteúdo para baixo.|
| _Up_|Opcional|**Variant**|O número de linhas pelas quais rolar o conteúdo para cima.|
| _ToRight_|Opcional|**Variant**|O número de colunas pelas quais rolar o conteúdo para a direita.|
| _ToLeft_|Opcional|**Variant**|O número de colunas pelas quais rolar o conteúdo para a esquerda.|

### Valor retornado

Variante


## Comentários

Se  _Down_ e _Up_ forem especificados, o conteúdo da janela será rolado pela diferença dos argumentos. Por exemplo, se _Down_ for 3 e _Up_ for 6, o conteúdo será rolado três linhas para cima.

Se  _ToLeft_ e _ToRight_ forem especificados, o conteúdo da janela será rolado pela diferença dos argumentos. Por exemplo, se _ToLeft_ for 3 e _ToRight_ for 6, o conteúdo será rolado três colunas para a direita.

Qualquer desses argumentos pode ser um número negativo.


## Exemplo

Este exemplo rola o conteúdo da janela ativa de Sheet1 para baixo três linhas.


```
Worksheets("Sheet1").Activate 
ActiveWindow.SmallScroll down:=3
```


## Ver também


#### Conceitos


[Objeto Window](8591b1ad-76f8-14e2-9120-406b65093f5a.md)
#### Outros recursos


[Membros do objeto Window](f11db427-24a4-041c-2fd5-03ce73ae6c16.md)