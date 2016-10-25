
# Estouro (Erro 6)

Ocorre um estouro quando você tenta fazer uma atribuição que excede as limitações do destino da atribuição. Esse erro tem as seguintes causas e soluções:



- O resultado de uma atribuição, um cálculo ou uma conversão de [tipo de dados](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) é muito grande para ser representado dentro do intervalo de valores permitidos para esse tipo de[variável](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).
    
    Atribua o valor a uma variável de um tipo que possa conter um intervalo de valores maior.
    
- Uma atribuição para uma [propriedade](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) excede o valor máximo que a propriedade pode aceitar.
    
    Verifique se a atribuição é adequada ao intervalo da propriedade para a qual ela é feita.
    
- Você tenta usar um número em um cálculo, e o número é imposto como um inteiro, mas o resultado é maior do que um inteiro. Por exemplo:
    
  ```
      Dim x As Long 
    x = 2000 * 365   ' Error: Overflow
  ```


    Para contornar a situação, digite o número da seguinte forma:
    


  ```
      Dim x As Long 
    x = 2000 * 365   ' Error: Overflow
  ```




  ```
      Dim x As Long 
    x = CLng(2000) * 365
  ```


    Para contornar a situação, digite o número da seguinte forma:
    


  ```
      Dim x As Long 
    x = CLng(2000) * 365
  ```


Para saber mais, selecione o item em questão e pressione F1 (no Windows) ou AJUDA (no Macintosh).
