
# ColumnWidth-Eigenschaft

Gibt die Breite aller Spalten im angegebenen Bereich zurück oder legt sie fest. Variant-Wert mit Lese-/Schreibzugriff.

 _expression_. **ColumnWidth**

 _Ausdruck_ ist erforderlich. Ein Ausdruck, der eines der Objekte in der Liste "Wird angewendet auf" zurückgibt.


## Hinweise

Eine Einheit der Spaltenbreite entspricht der Breite eines Zeichens im Format  **Normal**. Für proportionale Schriftarten wird die Breite des Zeichens 0 (Null) verwendet.

Besitzen alle Spalten in dem Bereich dieselbe Breite, so gibt die  **ColumnWidth** -Eigenschaft diese Breite zurück. Sind die Spalten unterschiedlich breit, so gibt die Eigenschaft **Null** zurück.


## Beispiel

In diesem Beispiel wird die Breite der Spalte A im Datenblatt verdoppelt.


```
With myChart.Application.DataSheet.Columns("A") 
 .ColumnWidth = .ColumnWidth * 2 
End With
```

