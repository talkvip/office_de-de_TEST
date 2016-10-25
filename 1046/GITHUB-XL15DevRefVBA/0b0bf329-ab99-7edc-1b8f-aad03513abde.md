
# Método Range.TextToColumns (Excel)

Analisa uma coluna de células que contêm texto em várias colunas.


## Sintaxe

 _expressão_. **TextToColumns**( ** _Destination_**, ** _DataType_**, ** _TextQualifier_**, ** _ConsecutiveDelimiter_**, ** _Tab_**, ** _Semicolon_**, ** _Comma_**, ** _Space_**, ** _Other_**, ** _OtherChar_**, ** _FieldInfo_**, ** _DecimalSeparator_**, ** _ThousandsSeparator_**, ** _TrailingMinusNumbers_** )

 _expressão_ Uma variável que representa um objeto **Range**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Destination_|Opcional|**Variant**|Um objeto  **Range** que especifica onde o Microsoft Excel colocará os resultados. Se o intervalo for maior do que uma única célula, a célula superior esquerda será usada.|
| _DataType_|Opcional|**[XlTextParsingType](71d76a41-c0b0-0b0f-27b5-7cac0d4c4ac4.md)**|O formato do texto a ser dividido em colunas.|
| _TextQualifier_|Opcional|**[XlTextQualifier](ba209892-9dea-84db-eafd-629c7ab0b20f.md)**|Especifica se as aspas deverão ser usadas ou não como o qualificador de texto e, se forem usadas, se serão simples ou duplas.|
| _ConsecutiveDelimiter_|Opcional|**Variant**|**True** para que o Microsoft Excel considere delimitadores consecutivos como um delimitador. O valor padrão é **False**.|
| _Tab_|Opcional|**Variant**|**True** para ter _DataType_ como **xlDelimited** e para ter o caractere de tabulação como um delimitador. O valor padrão é **False**.|
| _Semicolon_|Opcional|**Variant**|**True** para ter _DataType_ como **xlDelimited** e para ter o ponto-e-vírgula como um delimitador. O valor padrão é **False**.|
| _Comma_|Opcional|**Variant**|**True** para ter _DataType_ como **xlDelimited** e para ter a vírgula como um delimitador. O valor padrão é **False**.|
| _Space_|Opcional|**Variant**|**True** para ter _DataType_ como **xlDelimited** e para ter o caractere de espaço como um delimitador. O valor padrão é **False**.|
| _Other_|Opcional|**Variant**|**True** para ter _DataType_ como **xlDelimited** e para ter o caractere especificado pelo argumento _OtherChar_ como um delimitador. O valor padrão é **False**.|
| _OtherChar_|Opcional|**Variant**|(obrigatório se  _Other_ for **True** ). O caractere delimitador quando _Other_ for **True**. Se mais de um caractere for especificado, somente o primeiro caractere da cadeia de caracteres será usado; os caracteres restantes serão ignorados.|
| _FieldInfo_|Opcional|**Variant**|Uma matriz com informações de análise para colunas de dados individuais. A interpretação depende do valor de  _DataType_. Quando os dados forem delimitados, esse argumento será uma matriz de matrizes de dois elementos, com cada matriz de dois elementos especificando as opções de conversão para uma coluna em particular. O primeiro elemento é o número da coluna (baseada em 1) e o segundo elemento é uma das constantes de [xlColumnDataType](034f6011-c860-0887-9661-857821f630e4.md) que especificam como a coluna é analisada.|
| _DecimalSeparator_|Opcional|**Variant**|O separador decimal que o Microsoft Excel usa quando reconhece números. A configuração padrão é aquela do sistema.|
| _ThousandsSeparator_|Opcional|**Variant**|O separador de milhares que o Excel usa quando reconhece números. A configuração padrão é aquela do sistema.|
| _TrailingMinusNumbers_|Opcional|**Variant**|Números que começam com um sinal de menos.|

### Valor de retorno

Variante


## Comentários

A tabela a seguir mostra os resultados da importação de texto para o Excel para diversas configurações de importação. Os resultados numéricos são exibidos na coluna mais à direita.



|**Separador decimal do sistema**|**Separador de milhar do sistema**|**Valor do separador decimal**|**Valor do separador de milhares**|**Texto original**|**Valor da célula (tipo de dados)**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|Ponto|Vírgula|Vírgula|Ponto|123,123.45|123,123.45 (numérico)|
|Ponto|Vírgula|Vírgula|Vírgula|123,123.45|123,123.45 (texto)|
|Vírgula|Ponto|Vírgula|Ponto|123,123.45|123,123.45 (numérico)|
|Ponto|Vírgula|Ponto|Vírgula|123 123.45|123 123.45 (texto)|
|Ponto|Vírgula|Ponto|Espaço|123 123.45|123,123.45 (numérico)|

||
|:-----|
|**XlColumnDataType** pode ser uma dessas constantes **XlColumnDataType**.|
|**xlGeneralFormat**. Geral|
|**xlTextFormat**. Texto|
|**xlMDYFormat**. Data MDA|
|**xlDMYFormat**. Data DMA|
|**xlYMDFormat**. Data AMD|
|**xlMYDFormat**. Data MAD|
|**xlDYMFormat**. Data DAM|
|**xlYDMFormat**. Data ADM|
|**xlEMDFormat**. Data EMD|
|**xlSkipColumn**. Ignorar Coluna|
Você só poderá usar  **xlEMDFormat** se tiver instalado e selecionado o suporte ao idioma taiwanês. A constante **xlEMDFormat** especifica que datas da era taiwanesa estão sendo usadas.

Os especificadores de coluna podem estar em qualquer ordem. Se um dado especificador de coluna não estiver presente para uma determinada coluna nos dados de entrada, a coluna será analisada com a configuração  **xlGeneralFormat**. Este exemplo faz com que a terceira coluna seja ignorada, a primeira coluna seja analisada como texto e as colunas restantes nos dados de origem sejam analisadas com a configuração **xlGeneralFormat**.

 `Array(Array(3, 9), Array(1, 2))`

Se os dados de origem tiverem colunas de largura fixa, o primeiro elemento de cada matriz de dois elementos especificará a posição do caractere inicial na coluna (como um inteiro; 0 (zero) é o primeiro caractere). O segundo elemento da matriz de dois elementos especifica a opção de análise para a coluna como um número de 1 a 9, como listado acima.

O exemplo a seguir analisa duas colunas de um arquivo de largura fixa, com a primeira coluna desde o início da linha e se estendendo por 10 caracteres. A segunda coluna começa na posição 15 e prossegue até o fim da linha. Para evitar incluir os caracteres entre a posição 10 e a posição 15, o Microsoft Excel adiciona uma entrada de coluna ignorada.

 `Array(Array(0, 1), Array(10, 9), Array(15, 1))`


## Exemplo

Este exemplo converte o conteúdo da Área de Transferência, que contém uma tabela de texto delimitado por espaços, em colunas separadas na Planilha1. Você pode criar uma tabela simples delimitada no Bloco de Notas ou no WordPad (ou em outro editor de texto), copiar a tabela de texto para a Área de Transferência, alternar para o Microsoft Excel e executar esse exemplo.


```
Worksheets("Sheet1").Activate 
ActiveSheet.Paste 
Selection.TextToColumns DataType:=xlDelimited, _ 
 ConsecutiveDelimiter:=True, Space:=True
```


## Ver também


#### Conceitos


[Objeto Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Outros recursos


[Membros do objeto Range](4336bf81-1e63-7e44-1792-baf366a027a7.md)