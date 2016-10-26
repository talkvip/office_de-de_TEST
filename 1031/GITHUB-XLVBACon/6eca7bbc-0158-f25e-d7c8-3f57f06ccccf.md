
# ChartTitle-Objekt

Stellt den Titel des angegebenen Diagramms dar.


## Verwenden des ChartTitle-Objekts

Verwenden Sie die  **ChartTitle** -Eigenschaft, um das **ChartTitle** -Objekt zurückzugeben. Im folgenden Beispiel wird dem Diagramm ein Titel hinzugefügt.


```
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Hinweise

Das  **ChartTitle** -Objekt ist nicht vorhanden und kann nicht verwendet werden, wenn die **[HasTitle](9ecc48d3-fd86-e185-a69f-0676230b3194.md)** -Eigenschaft des Diagramms auf **True** festgelegt ist.

