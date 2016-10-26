
# MinimumScale-Eigenschaft

Gibt den Minimalwert auf der Größenachse zurück oder legt ihn fest.  **Double** -Wert mit Lese-/Schreibzugriff.


## Hinweise

Durch Festlegen dieser Eigenschaft wird die  **[MinimumScaleIsAuto](95ed7a2b-efda-b05a-da2e-789a166a97c8.md)** -Eigenschaft auf **false festgelegt**.


## Beispiel

In diesem Beispiel werden der niedrigste und der höchste Wert für die Größenachse festgelegt.


```
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```

