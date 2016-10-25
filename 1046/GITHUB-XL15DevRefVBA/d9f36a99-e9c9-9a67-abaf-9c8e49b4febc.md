
# Propriedade Application.DisplayAlerts (Excel)

 **True** se o Microsoft Excel exibir determinados alertas e mensagens durante a execução de uma macro. Ler/gravar **Boolean**.


## Sintaxe

 _expressão_. **DisplayAlerts**

 _expressão_ Uma variável que representa um objeto **Application**.


## Comentários

O valor padrão é  **True**. Defina essa propriedade como **False** para suprimir avisos e mensagens de alerta durante a execução de uma macro; quando uma mensagem exigirem uma resposta, o Microsoft Excel escolhe a resposta padrão.

Se você definir essa propriedade como  **False**, o Microsoft Excel definirá essa propriedade como **True** quando o código for finalizado, a menos que você esteja executando código de processo cruzado.




 **Observação**  Ao usar o método  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** para pastas de trabalho para substituir um arquivo existente, a caixa de diálogo **Confirmar Salvar como** terá como padrão **Não**, embora a resposta  **Sim** seja selecionada pelo Excel quando a propriedade **DisplayAlerts** estiver definida como **False**. A resposta **Sim** substitui o arquivo existente.Ao usar o método  **[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)** para pastas de trabalho para salvar uma pasta de trabalho que contém um projeto Visual Basic for Applications (VBA) no formato de arquivo do Excel 5.0/95, a caixa de diálogo **Microsoft Excel** tem um padrão **Sim**, embora a resposta  **Cancelar** seja selecionada pelo Excel quando a propriedade **DisplayAlerts** estiver definida como **False**. Não é possível salvar uma pasta de trabalho com um projeto VBA usando o formato de arquivo do Excel 5.0/95.


## Exemplo

Este exemplo fecha a pasta de trabalho Pasta1 e não solicita que o usuário salve as alterações. As alterações feitas na Pasta1.xls não são salvas.


```
Application.DisplayAlerts = False 
Workbooks("BOOK1.XLS").Close 
Application.DisplayAlerts = True
```

Este exemplo suprime a mensagem que, caso contrário, aparecerá quando você iniciar um canal DDE para um aplicativo que não esteja em execução.




```
Application.DisplayAlerts = False 
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DisplayAlerts = True 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber 
Application.DisplayAlerts = True
```


## Ver também


#### Conceitos


[Objeto Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Outros recursos


[Membros do objeto Application](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)