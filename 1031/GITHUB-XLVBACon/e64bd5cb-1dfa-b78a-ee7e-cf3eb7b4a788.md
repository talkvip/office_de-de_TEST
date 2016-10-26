
# VaryByCategories-Eigenschaft

 **True**, wenn Microsoft Graph jedem Datenpunkt eine andere Farbe oder ein anderes Muster zuweist. Das Diagramm darf nur eine Datenreihe enthalten. **Boolean** -Wert mit Lese-/Schreibzugriff.


## Beispiel

In diesem Beispiel wird jeder Datenpunktmarkierung in der ersten Diagrammgruppe eine andere Farbe oder ein anderes Muster zugewiesen. Es ist sinnvoll, das Beispiel auf ein 2D-Liniendiagramm anzuwenden, bei dem Datenpunktmarkierungen für eine Datenreihe vorhanden sind.


```
myChart.ChartGroups(1).VaryByCategories = True
```

