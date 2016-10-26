
# MajorUnitScale-Eigenschaft

Zurückgeben oder Festlegen der wichtigsten Einheitenwert des Skalierung für die Rubrikenachse, wenn die CategoryType-Eigenschaft auf XlTimeScale festgelegt ist. XlTimeUnit Lese-/Schreibzugriff.


||
|:-----|
|**XlTimeUnit** kann eine der folgenden **XlTimeUnit** -Konstanten sein.|
|**xlDays**|
|**xlMonths**|
|**xlYears**|

 _expression_. **MajorUnitScale**

 _Ausdruck_ ist erforderlich. Ein Ausdruck, der eines der Objekte in der Liste "Wird angewendet auf" zurückgibt.

## Beispiel

Dieses Beispiel legt für die Rubrikenachse die Verwendung einer Zeitskala sowie die Haupt- und Hilfseinheiten fest.


```
With myChart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .MajorUnit = 5 
 .MajorUnitScale = xlDays 
 .MinorUnit = 1 
 .MinorUnitScale = xlDays 
End With
```

