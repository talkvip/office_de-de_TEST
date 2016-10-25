
# Evento TextBox.Change (Access)

O evento  **Change** ocorre quando o conteúdo do controle especificado é alterado.


## Sintaxe

 _expressão_. **Change**

 _expressão_ Uma variável que representa um objeto **TextBox**.


## Comentários

São exemplos desse evento a entrada de um caractere diretamente na caixa de texto ou de combinação ou a alteração da configuração da propriedade  **Text** do controle usando uma macro ou o Visual Basic.


 **Observação**  A definição do valor de um controle usando uma macro ou o Visual Basic não dispara esse evento para o controle. É preciso digitar os dados diretamente no controle ou definir a propriedade  **Text** do controle.

Para executar uma macro ou um procedimento de evento quando esse evento ocorre, defina a propriedade  **OnChange** com o nome da macro ou como [Procedimento do Evento].

Você pode coordenar a exibição dos dados entre os controles ao executar uma macro ou um procedimento de evento quando um evento Change ocorre. Você pode também exibir dados ou uma fórmula em um controle e os resultados em um outro controle.

O evento Change não ocorre quando um valor é alterado em um controle calculado.


 **Observação**  Um evento Change pode causar um evento em cascata. Isso ocorre quando uma macro ou um procedimento de evento, executado em resposta ao evento Change do controle, altera o conteúdo do controle. Por exemplo, alterar uma configuração de propriedade que determina o valor do controle, como a propriedade  **Text** de uma caixa de texto. Para evitar um evento em cascata:


- Se possível, evite anexar uma macro ou procedimento de evento Change a um controle que altere o conteúdo deste.
    
- Evite criar dois ou mais controles tendo eventos Change que afetem um ao outro  por exemplo, duas caixas de texto que atualizem uma a outra.
    
Alterar os dados de uma caixa de texto ou caixa de combinação por meio do teclado faz com que eventos de teclado ocorram além dos eventos de controle, como o evento Change. Por exemplo, se você se mover para um novo registro e digitar um caractere ANSI em uma caixa de texto no registro, os eventos a seguir ocorrerão nesta ordem:

 **KeyDown** → **KeyPress** → **BeforeInsert** → **Change** → **KeyUp**

Os eventos  **BeforeUpdate** e **AfterUpdate** para o controle de caixa de texto ou caixa de combinação ocorrem depois de inserir os dados novos ou alterados no controle e alternar para outro controle (ou clicar em **Salvar Registro** no menu **Registros** ). Portanto, após todos os eventos Change para o controle.

Em caixas de combinação para as quais a propriedade  **LimitToList** está definida como Sim, o evento **NotInList** ocorre após inserir um valor que não está na lista e tentar mover para outro controle ou salvar o registro. Isso ocorre depois de todos os eventos Change para a caixa de combinação. Nesse caso, os eventos BeforeUpdate e AfterUpdate para a caixa de combinação não ocorrem porque o Microsoft Access não aceita um valor que não está na lista.


## Ver também


#### Conceitos


[Objeto TextBox](d74fbe9a-0d40-7d28-956f-a2bfd0cfee45.md)
#### Outros recursos


[Membros do objeto TextBox](bb55abbc-902e-fc2d-bdff-063c55426cd0.md)