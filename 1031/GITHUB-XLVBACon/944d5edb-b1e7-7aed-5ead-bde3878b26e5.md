
# Point-Objekt

Stellt einen einzelnen Punkt einer Datenreihe im angegebenen Diagramm dar.  **Point** -Objekt ist ein Mitglied der **[Points](b41c8f08-880e-1f4a-0456-3f77c0741bc6.md)** -Auflistung, die alle Datenpunkte in der angegebenen Datenreihe enthält.


## Verwenden des Point-Objekts

Verwenden Sie  **Punkt** ( _Index_ ), wobei _Index_ den Punkt Indexnummer ist, zur Rückgabe eines einzelnen Objekts **zeigen**. Punkt von links nach rechts nummeriert werden in der Datenreihe. `Points(1)` ist die am weitesten links und ist die am weitesten links und `Points(Points.Count)` ist der äußersten rechten. Im folgende Beispiel wird die Art der Markierung für den dritten Punkt der ersten Datenreihe. Für das Beispiel funktioniert muss Datenreihe 2D-Diagramme Linien, Punkt (XY) oder Netzdiagramme Datenreihe.


```
myChart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
```

