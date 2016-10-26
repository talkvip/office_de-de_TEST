
# ChartGroups-Auflistung (Excel)

Eine Auflistung aller  **[ChartGroup](8a485a8c-e181-a039-60b9-a02c2c89b26e.md)** -Objekte im angegebenen Diagramm. Jedes **ChartGroup** -Objekt stellt eine oder mehrere Datenreihen mit demselben Format in ein Diagramm gezeichnet. Ein Diagramm enthält eine oder mehrere Diagrammgruppen, jede Diagrammgruppe enthält eine oder mehrere Datenreihen und jede Datenreihe enthält einen oder mehrere Punkte. Beispielsweise kann ein einzelnes Diagramm sowohl eine Liniendiagrammgruppe, mit der Liniendiagramm gezeichneten enthält und eine Balkendiagrammgruppe, die mit dem Format des Balkendiagramms gezeichneten enthält enthalten.


## Verwenden der ChartGroups-Auflistung

Verwenden Sie die  **ChartGroups** -Methode, um die **ChartGroups** -Auflistung zurückzugeben. Im folgenden Beispiel wird die Anzahl der Diagrammgruppen im Diagramm angezeigt.


```
MsgBox myChart.ChartGroups.Count
```

Verwenden Sie  **ChartGroups** ( _Index_ ), wobei _Index_ die Indexnummer der Diagrammgruppe ist, um ein einzelnes **ChartGroup** -Objekt zurückzugeben. Im folgenden Beispiel werden der ersten Diagrammgruppe im Diagramm Bezugslinien hinzugefügt.




```
myChart.ChartGroups(1).HasDropLines = True
```

Da sich die Indexnummer einer bestimmten Diagrammgruppe ändern kann, wenn das für die Gruppe verwendete Diagrammformat geändert wird, kann es einfacher sein, eine der genannten Schnellzugriffsmethoden für Diagrammgruppen zu verwenden, um eine bestimmte Diagrammgruppe zurückzugeben. Die  **PieGroups** -Methode gibt die Auflistung von Kreisdiagrammgruppen in einem Diagramm zurück, die **LineGroups** -Methode gibt die Auflistung aller Liniendiagrammgruppen zurück usw. Sie können jede dieser Methoden mit einer Indexnummer verwenden, um ein einzelnes **ChartGroup** -Objekt zurückzugeben. Ohne Angabe einer Indexnummer wird eine **ChartGroups** -Auflistung zurückgegeben. Folgende Methoden sind für Diagrammgruppen verfügbar:


-  **[AreaGroups](ec2a4a28-2f10-4f4f-bd91-642bf1b8ebe2.md)** -Methode
    
-  **[BarGroups](a00e484e-05ec-2eaa-cc33-05b77a4af0b5.md)** -Methode
    
-  **[ColumnGroups](dcb4d7e0-ce56-46d9-35d9-d9653bbb6f97.md)** -Methode
    
-  **[DoughnutGroups](41ca4213-c17b-7bba-c357-7ba65fd55d39.md)** -Methode
    
-  **[LineGroups](3a8083b5-8b71-e28b-c775-6be50544d6b2.md)** -Methode
    
-  **[PieGroups](f7fd5497-f7a0-6c28-1a59-9e6f37a0885e.md)** -Methode
    
