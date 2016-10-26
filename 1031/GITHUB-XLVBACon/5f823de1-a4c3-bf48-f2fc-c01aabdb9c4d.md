
# DataLabel-Objekt

Stellt die Beschriftung für den angegebenen Punkt oder Trendlinie in einem Diagramm dar. Für eine Datenreihe ist das  **DataLabel** -Objekt ein Mitglied der **[DataLabels](597c7269-71ed-5dcc-af6b-34dc908e9d58.md)** -Auflistung, die ein **DataLabel** -Objekt für jeden Punkt enthält. Bei einer Datenreihe ohne definierbare Punkte (beispielsweise eine Bereichsdatenreihe) enthält die **DataLabels** -Auflistung ein einzelnes **DataLabel** -Objekt.


## Verwenden des DataLabel-Objekts

Verwenden Sie  **DataLabels** ( _Index_ ), wobei _Index_ die Indexnummer der Datenbeschriftung ist, um ein einzelnes **DataLabel** -Objekt zurückzugeben. Im folgenden Beispiel wird das Zahlenformat für die fünfte Datenbeschriftung der ersten Datenreihe im Diagramm festgelegt.


```
myChart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

Verwenden Sie die  **DataLabel** -Eigenschaft, um das **DataLabel** -Objekt für einen einzelnen Punkt zurückzugeben. Im folgenden Beispiel wird für den zweiten Punkt der ersten Datenreihe im Diagramm die Datenbeschriftung eingeschaltet und ihr der Text "Samstag" zugewiesen.




```
With myChart 
 With .SeriesCollection(1).Points(2) 
 .HasDataLabel = True 
 .DataLabel.Text = "Saturday" 
 End With 
End With
```

Bei einer Trendlinie gibt die  **DataLabel** -Eigenschaft den mit der Trendlinie angezeigten Text zurück. Dabei kann es sich um die Gleichung, den R-Quadratwert oder um beide handeln (wenn beide angezeigt werden). Im folgenden Beispiel wird festgelegt, dass der Text der Trendlinie nur die Gleichung anzeigt. Danach wird der Text der Datenbeschriftung in die Zelle A1 auf dem Datenblatt positioniert.




```
With myChart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 x = .DataLabel.Text 
End With 
With myChart.Application.DataSheet 
 .Range("A1").Value = x 
End With
```

