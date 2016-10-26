
# Verwenden von Ereignissen mit dem Application-Objekt

Bevor Sie ein Ereignis mit dem  **Application** -Objekt verwenden können, müssen Sie ein Klassenmodul erstellen und für Ereignisse ein Objekt vom Typ **Application** deklarieren. Nehmen Sie beispielsweise an, dass ein neues Klassenmodul erstellt wird, das **EventClassModule** heißt. Das neue Klassenmodul enthält den folgenden Code:


```
Public WithEvents App As Application
```


Nachdem das neue Objekt mit Ereignissen deklariert wurde, wird es im Klassenmodul im Dropdownlistenfeld  **Objekt** angezeigt, und Sie können Ereignisprozeduren für das neue Objekt schreiben. (Wenn Sie das neue Objekt im Feld **Objekt** auswählen, werden die gültigen Ereignisse für dieses Objekt im Dropdownlistenfeld **Prozedur** angezeigt.)

Bevor die Prozeduren ausgeführt werden könne, müssen Sie jedoch das deklarierte Objekt im Klassenmodul mit dem  **Application** -Objekt verbinden. Sie können dazu in allen Modulen den folgenden Code verwenden.

## Beispiel


```
Dim X As New EventClassModule 
 
Sub InitializeApp() 
 Set X.App = Application 
End Sub
```

Nachdem die  **InitializeApp** -Prozedur ausgeführt wurde, zeigt das **App** -Objekt im Klassenmodul auf das **Application** -Objekt von Microsoft Excel, und die Ereignisprozeduren im Klassenmodul werden ausgeführt, wenn die Ereignisse auftreten.

