
# MarkerBackgroundColor-Eigenschaft

Gibt die Hintergrundfarbe für die Markierungen als RGB-Wert zurück oder legt sie fest. Gilt nur für Linien-, Punkt- und Netzdiagramme.  **Long** -Wert mit Lese-/Schreibzugriff.


## Beispiel

Dieses Beispiel legt den Markierungshintergrund und die Vordergrundfarben für den zweiten Punkt der Reihe 1 fest.


```
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```

