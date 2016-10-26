
# BarShape-Eigenschaft

Gibt oder die mit einem angegebenen 3D-Balkendiagramm oder Säulendiagramm verwendete Form festgelegt. XlBarShape Lese-/Schreibzugriff.


||
|:-----|
|**XlBarShape** kann eine der folgenden **XlBarShape** -Konstanten sein.|
|**xlConeToMax**|
|**xlCylinder**|
|**xlPyramidToPoint**|
|**xlBox**|
|**xlConeToPoint**|
|**xlPyramidToMax**|

 _expression_. **BarShape**

 _Ausdruck_ ist erforderlich. Ein Ausdruck, der eines der Objekte in der Liste "Wird angewendet auf" zurückgibt.

## Beispiel

In diesem Beispiel wird die Form festgelegt, die mit der ersten Datenreihe im Diagramm verwendet wird.


```
myChart.SeriesCollection(1).BarShape = xlConeToPoint
```

