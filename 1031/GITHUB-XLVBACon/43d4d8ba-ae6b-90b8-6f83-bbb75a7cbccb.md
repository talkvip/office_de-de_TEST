
# Cells-Eigenschaft

Gibt ein Range-Objekt zurück, das die Zellen des angegebenen Bereichs darstellt, wie er auf das Range-Objekt Anwendung findet. Außerdem wird ein Range-Objekt zurückgegeben, das alle Zellen des Datenblattes (nicht nur die aktuell verwendeten) darstellt, so wie es auf das DataSheet-Objekt angewendet wird. Schreibgeschütztes Range-Objekt.

 _expression_. **Cells**

 _Ausdruck_ Erforderlich. Ein Ausdruck, der in der Liste gilt für ein Objekt zurückgibt.


## Beispiel

In diesem Beispiel wird im Datenblatt die Formel in der Zelle A1 gelöscht. Beachten Sie, dass es sich im Datenblatt bei der Spalte A um die zweite Spalte und bei der Zeile 1 um die zweite Zeile handelt.


```
myChart.Application.DataSheet.Cells(2,2).ClearContents
```

In diesem Beispiel werden im Datenblatt die Zellen A1:A13 in einer Schleife bearbeitet. Besitzen dabei Zellen einen Wert, der kleiner ist als 0,001, wird er durch 0 (Null) ersetzt.




```
Set mySheet = myChart.Application.DataSheet 
For rwIndex = 2 to 4 
 For colIndex = 2 to 10 
 If mySheet.Cells(rwIndex, colIndex) < .001 Then 
 mySheet.Cells(rwIndex, colIndex).Value = 0 
 End If 
 Next colIndex 
Next rwIndex
```

