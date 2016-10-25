
# Método TextBox.SetFocus (Access)

O método  **SetFocus** move o foco para o formulário especificado, para o controle especificado no formulário ativo ou para o campo especificado na folha de dados ativa.


## Sintaxe

 _expressão_. **SetFocus**

 _expressão_ Uma variável que representa um objeto **TextBox**.


### Valor retornado

Nada


## Comentários

Você poderá usar o método  **SetFocus** quando quiser que um campo ou um controle em particular tenha o foco de forma que todas as entradas de usuário sejam direcionadas para esse objeto.

Para ler algumas das propriedades de um controle, será necessário garantir que o controle tenha o foco. Por exemplo, uma caixa de texto deverá ter o foco antes que você possa ler o  **Texto** dela

Outras propriedades só poderão ser definidas quando um controle não tiver o foco. Por exemplo, não é possível definir as propriedades  **Visible** ou **Enabled** de um controle como **False** (0) quando o controle tiver o foco.

Você também pode usar o método  **SetFocus** para navegar em um formulário de acordo com determinadas condições. Por exemplo, se o usuário selecionar **Not applicable** para o primeiro de um conjunto de perguntas em um formulário que é um questionário, seu código do Visual Basic poderia então ignorar automaticamente as perguntas desse conjunto e mover o foco para o primeiro controle no próximo conjunto de perguntas.

Você só pode mover o foco para um controle ou formulário visível. Um formulário e controles em um formulário não estarão visíveis até a conclusão do evento  **Load** do formulário. Dessa forma, se você usar o método **SetFocus** no evento Carregamento de um formulário para mover o foco para o formulário, deverá usar o método **Repaint** antes do método **SetFocus**.

Não é possível mover o foco para um controle caso a sua propriedade  **Enabled** esteja definida como **False**. Você deve definir a propriedade **Enabled** de um controle como **True** (-1) antes de poder mover o foco para esse controle. No entanto, você pode mover o foco para um controle caso a sua propriedade **Locked** esteja definida como **True**.

Se um formulário contiver controles para os quais a propriedade  **Enabled** esteja definida como **True**, você não poderá mover o foco para o próprio formulário. Você só poderá mover o foco para controles no formulário. Nesse caso, se você tentar usar **SetFocus** para mover o foco para um formulário, o foco será definido para o controle no formulário que recebeu o foco pela última vez.

Você pode usar o método  **SetFocus** para mover o foco para um subformulário, que é um tipo de controle. Você também pode mover o foco para um controle em um subformulário usando o método **SetFocus** duas vezes, movendo o foco primeiro para o subformulário e depois para o controle no subformulário.


## Exemplo

O exemplo a seguir usa o método  **SetFocus** para mover o foco para uma caixa de texto EmployeeID em um formulário Funcionários:


```
Forms!Employees!EmployeeID.SetFocus
```


## Ver também


#### Conceitos


[Objeto TextBox](d74fbe9a-0d40-7d28-956f-a2bfd0cfee45.md)
#### Outros recursos


[Membros do objeto TextBox](bb55abbc-902e-fc2d-bdff-063c55426cd0.md)