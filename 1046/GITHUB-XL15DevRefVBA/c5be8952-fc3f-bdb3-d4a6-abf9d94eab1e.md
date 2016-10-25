
# Propriedade Range.Formula (Excel)

Retorna ou define um valor  **Variant** que representa a fórmula do objeto na notação estilo A1 e na linguagem de macro.


## Sintaxe

 _expressão_. **Formula**

 _expressão_ Uma variável que representa um objeto **Range**.


## Comentários

Esta propriedade não está disponível para fontes de dados OLAP.

Se a célula contiver uma constante, essa propriedade retornará a constante. Se a célula estiver vazia, essa propriedade retornará uma cadeia de caracteres vazia. Se a célula contiver uma fórmula, a propriedade  **Formula** retornará a fórmula como uma cadeia de caracteres no mesmo formato em que seria exibida na barra de fórmulas (incluindo o sinal de mais (=)).

Se você definir o valor ou a fórmula de uma célula como uma data, O Microsoft Excel verificará se a célula já está formatada com um dos formatos de número de data e hora. Caso contrário, o Microsoft Excel mudará o formato de número para o formato de número de data curta padrão.

Se o intervalo tiver uma ou duas dimensões, você poderá definir a fórmula como uma matriz do Visual Basic com as mesmas dimensões. Da mesma forma, é possível colocar a fórmula em uma matriz do Visual Basic.

A configuração da fórmula para um intervalo de várias células preenche todas as células do intervalo com a fórmula.


## Exemplo

O código a seguir define a fórmula para a célula A1 na Planilha1.


```
Worksheets("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
```



 **Código de exemplo fornecido por:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)

O exemplo de código a seguir define a fórmula para a célula A1 na Planilha1 para exibir a data de hoje.




```
Sub InsertTodaysDate() 
    ' This macro will put today's date in cell A1 on Sheet1 
    Sheets("Sheet1").Select 
    Range("A1").Select 
    Selection.Formula = "=text(now(),""mmm dd yyyy"")" 
    Selection.Columns.AutoFit 
End Sub
```


## Sobre o colaborador
<a name="AboutContributor"> </a>

MVP Bill Jelen é o autor de mais de duas dezenas de livros sobre o Microsoft Excel. Ele é um convidado regular do TechTV com Leo Laporte e é o anfitrião de MrExcel.com, que inclui mais de 300.000 perguntas e respostas sobre Excel.


## Ver também
<a name="AboutContributor"> </a>


#### Conceitos


[Objeto Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Outros recursos


[Membros do objeto Range](4336bf81-1e63-7e44-1792-baf366a027a7.md)