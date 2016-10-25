
# Erro definido pelo aplicativo ou pelo objeto

Esta mensagem é exibida quando um erro gerado através do método  **Raise** ou da declaração **Error** não corresponde a um erro definido pelo Visual Basic for Applications. Ela também é retornada pela função **Error** para[argumentos](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) que não correspondam a erros definidos pelo Visual Basic for Applications. Apesar de poder ser um erro definido por você, ou um que seja definido por um objeto, incluindo os[aplicativos host](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) como o Microsoft Excel, Visual Basic, etc. Por exemplo, os formulários do Visual Basic geram erros relacionados a formulários que não podem ser gerados a partir do código especificando um número como um argumento para o método **Raise** ou a declaração **Error**. Esta mensagem tem as seguintes causas e soluções:



- O seu aplicativo executou uma declaração  **Err.Raise** _n_ ou **Error** _n_, mas o número _n_ não foi definido pelo Visual Basic for Applications. Se esta era a intenção, você deve usar **Err.Raise** e especificar argumentos adicionais para que um usuário final possa entender a natureza do erro. Por exemplo, você pode incluir uma cadeia de caracteres de descrição, fonte e informações de ajuda. Para regenerar um erro interceptado, esta abordagem funcionará se você não executar **Err.Clear** antes de regenerar o erro. Se executar **Err.Clear** primeiro, deverá preencher os argumentos adicionais ao método **Raise**. Consulte o contexto em que o erro ocorreu e certifique-se de que está regenerando o mesmo erro.
    
- Ao acessar objetos em outros aplicativos, pode ter sido propagado de volta para o seu programa um erro que não possa ser mapeado para um erro do Visual Basic.
    
  ```
  For index = 1 to 500 
Debug.Print Error$(index) 
Next index 

  ```


    Consulte a documentação para ver os objetos que você possa ter acessado. A propriedade  **Source** do objeto **Err** deve conter a ID programática do aplicativo ou objeto que gerou o erro. Para entender o contexto de um erro retornado por um objeto, você poderá usar o constructo **On Error Resume Next** no código que acessa objetos, em vez da sintaxe de _linha_ **On Error GoTo**.
    
     **Observação**  No passado, os programadores usavam frequentemente um loop para imprimir uma lista de todas as cadeias de caracteres de mensagens de erro interceptáveis. Normalmente, isso era feito com um código como o seguinte:



  ```
  For index = 1 to 500
Debug.Print Error$(index)
Next index

  ```


     **Observação**  Esse código ainda lista todas as mensagens de erro do Visual Basic for Applications, mas exibe "Erro definido pelo aplicativo ou pelo objeto" para erros definidos pelo host, como os que se referem a formulários, controles, etc. no Visual Basic. Muitos desses erros são [erros de execução](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) interceptáveis. Você pode usar a caixa de diálogo **Pesquisar** da Ajuda para encontrar a lista de erros interceptáveis do seu aplicativo host. Clique em **Pesquisar**, digite **Interceptável** na primeira caixa de texto e clique em **Mostrar Tópicos**. Selecione **Erros Interceptáveis** na caixa de lista inferior e clique em **Ir Para**.

Para saber mais, selecione o item em questão e pressione F1 (no Windows) ou AJUDA (no Macintosh).
