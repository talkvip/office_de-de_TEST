
# Markieren und Aktivieren von Zellen

Mit Microsoft Excel markieren Sie gewöhnlich eine Zelle bzw. Zellen und führen dann eine Aktion aus, wie z. B. das Formatieren der Zellen oder das Eingeben von Werten in diese Zellen. In Visual Basic ist es in der Regel nicht notwendig, Zellen vor einer Änderung zu markieren.

Wenn Sie in Visual Basic z. B. eine Formel in Zelle D6 eingeben, brauchen Sie den Bereich D6 nicht zu markieren. Sie müssen nur das  **Range** -Objekt für die betreffende Zelle zurückgeben und dann für die gewünschte Formel die **Formula** -Eigenschaft festlegen, wie im folgenden Beispiel gezeigt wird.



```VB.net
Sub EnterFormula() 
    Worksheets("Sheet1").Range("D6").Formula = "=SUM(D2:D5)" 
End Sub
```

Weitere Informationen und Beispiele zur Verwendung anderer Methoden, Zellen ohne vorherige Markierung zu steuern, finden Sie unter [Vorgehensweise: Auf Zellen und Bereiche verweisen](a16caa8d-21c9-ff33-347b-ce671248a92d.md).

## Verwenden der Select-Methode und der Selection-Eigenschaft

Die  **Select** -Methode aktiviert Blätter und darin enthaltene Objekte; die **Selection** -Eigenschaft gibt ein Objekt zurück, das für die aktuelle Markierung im aktiven Blatt der aktiven Arbeitsmappe steht. Vor einem erfolgreichen Verwenden der **Selection** -Eigenschaft müssen Sie eine Arbeitsmappe aktivieren, ein Blatt aktivieren bzw. auswählen und anschließend mit der **Select** -Methode einen Bereich (oder ein anderes Objekt) markieren.

Der Makro-Recorder erstellt häufig Makros, die die  **Select** -Methode und die **Selection** -Eigenschaft verwenden. Die folgende **Sub** -Prozedur wurde mit dem Makro-Recorder erstellt und veranschaulicht, wie **Select** und **Selection** gemeinsam angewendet werden.




```VB.net
Sub Macro1() 
    Sheets("Sheet1").Select 
    Range("A1").Select 
    ActiveCell.FormulaR1C1 = "Name" 
    Range("B1").Select 
    ActiveCell.FormulaR1C1 = "Address" 
    Range("A1:B1").Select 
    Selection.Font.Bold = True 
End Sub
```

Im folgenden Beispiel wird die gleiche Aufgabe ohne vorheriges Aktivieren oder Markieren des Arbeitsblatts bzw. der Zellen ausgeführt.




```VB.net
Sub Labels() 
    With Worksheets("Sheet1") 
        .Range("A1") = "Name" 
        .Range("B1") = "Address" 
        .Range("A1:B1").Font.Bold = True 
    End With 
End Sub
```


## Markieren von Zellen im aktiven Arbeitsblatt

Wenn Sie zur Markierung von Zellen die  **Select** -Methode verwenden, bedenken Sie, dass **Select** nur im aktiven Arbeitsblatt funktioniert. Beim Ausführen Ihrer **Sub** -Prozedur vom Modul aus funktioniert die **Select** -Methode nur dann, wenn die Prozedur das Arbeitsblatt vor dem Verwenden der **Select** -Methode aktiviert. Als Beispiel kopiert die folgende Prozedur eine Zeile von Sheet1 in Sheet2 der aktiven Arbeitsmappe.


```VB.net
Sub CopyRow() 
    Worksheets("Sheet1").Rows(1).Copy 
    Worksheets("Sheet2").Select 
    Worksheets("Sheet2").Rows(1).Select 
    Worksheets("Sheet2").Paste 
End Sub
```


## Aktivieren einer Zelle innerhalb einer Markierung

Sie können zur Aktivierung einer Zelle innerhalb einer Markierung die  **Activate** -Methode verwenden. Selbst wenn ein Zellbereich markiert ist, kann immer nur eine Zelle aktiv sein. Die folgende Prozedur markiert einen Bereich und aktiviert danach eine Zelle innerhalb des Bereichs, ohne die Markierung zu ändern.


```VB.net
Sub MakeActive() 
    Worksheets("Sheet1").Activate 
    Range("A1:D4").Select 
    Range("B2").Activate 
End Sub
```

