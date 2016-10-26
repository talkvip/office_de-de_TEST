
# CapitalizeNamesOfDays-Eigenschaft

True, wenn die Anfangsbuchstaben von Wochentagsnamen automatisch großgeschrieben werden. Boolean-Wert mit Lese-/Schreibzugriff.

 _expression_. **CapitalizeNamesOfDays**

 _Ausdruck_ ist erforderlich. Ein Ausdruck, der eines der Objekte in der Liste "Wird angewendet auf" zurückgibt.


## Beispiel

In diesem Beispiel wird Microsoft Graph so eingestellt, dass der Anfangsbuchstabe der Wochentagsnamen großgeschrieben wird.


```
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```

