
# Evento QueryClose



Ocorre antes que um  **UserForm** feche.
 **Syntax**
 **Private Sub UserForm_QueryClose(** _cancel_ **As Integer**, _closemode_ **As Integer)**
A sintaxe do evento  **QueryClose** tem estas partes:


|**Parte**|**Descrição**|
|:-----|:-----|
| _cancelar_|Um número inteiro. A configuração desse [argumento](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) como qualquer valor diferente de 0 para o evento QueryClose em todos os formulários de usuário carregados e impede que o **UserForm** e o aplicativo fechem.|
| _closemode_|Um valor ou [constante](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) que indica a causa do evento QueryClose.|
 **Return Values**
O argumento  _closemode_ retorna os seguintes valores:


|**Constante**|**Valor**|**Descrição**|
|:-----|:-----|:-----|
|**vbFormControlMenu**|0|O usuário escolheu o comando  **Close** do menu **Control** no **UserForm**.|
|**vbFormCode**|1|A instrução  **Unload** é invocada do código.|
|**vbAppWindows**|2|A sessão operacional atual do Windows está terminando.|
|**vbAppTaskManager**|3|O  **Task Manager** do Windows está fechando o aplicativo.|
Estas constantes estão listadas na [biblioteca de objetos](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) do Visual Basic for Applications no[Pesquisador de Objetos](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md). Observe que  **vbFormMDIForm** também é especificado no **Object Browser**, mas ainda não tem suporte.
 **Remarks**
Este evento é tipicamente usado para garantir que não haja tarefas inacabadas nos formulários do usuário incluídas em um aplicativo antes que o aplicativo seja fechado. Por exemplo, se um usuário não tiver salvado dados novos em qualquer  **UserForm**, o aplicativo poderá solicitar que o usuário salve os dados.
Quando um aplicativo fechar, você poderá usar o procedimento de eventos  **QueryClose** para definir a propriedade **Cancel** como **True**, parando o processo de fechamento.

## Exemplo

O código a seguir força o usuário a clicar na área de cliente do  **UserForm** para fechá-lo. Se o usuário tentar usar a caixa **Close** na barra de título, o parâmetro _Cancel_ será definido com um valor diferente de zero, impedindo o término. Entretanto, se o usuário tiver clicado na área do cliente, _CloseMode_ terá o valor 1 e `Unload Me` será executado.


```
Private Sub UserForm_Activate()
    UserForm1.Caption = "You must Click me to kill me!"
End Sub

Private Sub UserForm_Click()
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
    UserForm1.Caption = "The Close box won't work! Click me!"
End Sub

```

