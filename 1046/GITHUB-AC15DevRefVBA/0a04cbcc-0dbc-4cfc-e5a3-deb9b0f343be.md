
# Método Application.DLast (Access)

Você pode usar a função  **DLast** para retornar um registro aleatório de um determinado campo em uma tabela ou consulta, quando precisar apenas de um valor qualquer desse campo.


## Sintaxe

 _expressão_. **DLast**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expressão_ Uma variável que representa um objeto **Application**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Expr_|Obrigatório|**String**|Uma expressão que identifica o campo do qual você deseja encontrar os valores mínimo e máximo. Pode ser uma expressão de cadeia de caracteres que identifica um campo em uma tabela ou consulta ou pode ser uma expressão que execute um [cálculo com os dados desse campo](http://msdn.microsoft.com/library/73c27d1c-0a3c-03e4-c17c-337133d7b316%28Office.15%29.aspx). Em  _expr_, inclua o nome de um campo em uma tabela, um controle em um formulário, uma constante ou uma função. Se  _expr_ incluir uma função, poderá ser interna ou definida pelo usuário, mas não um outro domínio agregado ou função SQL agregada.|
| _Domain_|Obrigatório|**String**|Uma expressão formada por cadeia de caracteres que identifica o conjunto de registros que constitui o domínio.|
| _Criteria_|Opcional|**Variant**|Uma expressão de cadeia de caracteres opcional utilizada para restringir o intervalo de dados no qual a função  **DLast** é executada. Por exemplo, _criteria_ costuma ser equivalente à cláusula WHERE em uma expressão SQL, sem o termo WHERE. Se _criteria_ for omitido, a função **DLast** avaliará _expr_ em relação ao domínio inteiro. Todos os campos incluídos em _criteria_ também deverão ser campos em _domain_; caso contrário, a função  **DLast** retornará **Null**.|

### Valor retornado

Variante


## Comentários




 **Observação**  Para retornar o primeiro ou o último registro de um conjunto de registros (um domínio), crie uma consulta classificada como crescente ou decrescente e defina a propriedade  **TopValues** como 1. No Visual Basic, você pode também criar um objeto **Recordset** do ADO e usar o método **MoveFirst** ou **MoveLast** para retornar o primeiro ou o último registro de um conjunto de registros.


## Exemplo



Os exemplos a seguir mostram como usar os diversos tipos de critérios com a função  **DLast**.

 **Código de exemplo fornecido por:**
![Ícone de Membro da Comunidade](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) A comunidade[UtterAccess](http://www.utteraccess.com)




```
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DLast("[FieldName]", "TableName", "[Criteria] = " &amp; Forms!FormName!ControlName)

    ' Strings
    variable = DLast("[FieldName]", "TableName", "[Criteria] = '" &amp; Forms!FormName!ControlName &amp; "'")

    ' Dates
    variable = DLast("[FieldName]", "TableName", "[Criteria] = #" &amp; Forms!FormName!ControlName &amp; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DLast("[FieldName]", "TableName", "[Criteria1] = " &amp; Forms![FormName]![Control1] _
             &amp; " AND [Criteria2] = '" &amp; Forms![FormName]![Control2] &amp; "'" _
            &amp; " AND [Criteria3] =#" &amp; Forms![FormName]![Control3] &amp; "#")
    
    ' Use two fields from a single record.
    variable = DLast("[LastName] &amp; ', ' &amp; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DLast("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DLast("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## Sobre os colaboradores
<a name="AboutContributors"> </a>

UtterAccess é o fórum principal de wiki e a Ajuda do Microsoft Access. clique aqui para ingressar.


## Ver também
<a name="AboutContributors"> </a>


#### Conceitos


[Objeto Application](aefb0713-97e6-e2c7-e530-8fd2e1316a55.md)
#### Outros recursos


[Membros do objeto Application](3ab5276c-d52a-72a9-244c-ec92ead48811.md)