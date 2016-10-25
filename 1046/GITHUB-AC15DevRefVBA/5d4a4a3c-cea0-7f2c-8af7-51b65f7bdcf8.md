
# Método DoCmd.RunCommand (Access)

O método  **RunCommand** executa um comando interno.


## Sintaxe

 _expressão_. **RunCommand**( ** _Command_** )

 _expressão_ Uma variável que representa um objeto **DoCmd**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Command_|Obrigatório|**AcCommand**|Uma constante  **[AcCommand](a78f91cc-3b40-5f45-c737-4d3abb2e979f.md)** que especifica o comando a ser executado.|

## Comentários

Cada comando de menu ou barra de ferramentas no Microsoft Access tem uma constante associada que pode ser usada com o método  **RunCommand** para executar esse comando do Visual Basic.

Não é possível usar o método  **RunCommand** para executar um comando em um menu ou barra de ferramentas personalizados. Ele só pode ser usado com menus e barras de ferramentas internos.

O método  **RunCommand** substitui o método **DoMenuItem** do objeto **DoCmd**.


## Ver também


#### Conceitos


[Objeto DoCmd](3ce44cca-9979-0a1e-9787-079a52ce528f.md)
#### Outros recursos


[Membros do objeto DoCmd](3e7ade9e-86e4-0751-188b-5d31c9101651.md)