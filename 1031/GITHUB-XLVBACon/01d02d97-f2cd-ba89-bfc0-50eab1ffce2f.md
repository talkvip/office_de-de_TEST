
# SolverLoad-Funktion

Lädt die Parameter für ein bestehendes Solvermodell, die in dem Arbeitsblatt gespeichert wurden.


 **Hinweis**  Das Solver-Add-In ist nicht standardmäßig aktiviert. Bevor Sie diese Funktion verwenden können, muss das Solver-Add-In aktiviert und installiert werden. Informationen hierzu finden Sie unter [Verwenden der Solver VBA-Funktionen](37d0aa49-2e5c-5efe-1c69-b5168af1f231.md). Nach der Installation des Solver-Add-Ins müssen Sie einen Verweis auf das Solver-Add-In erstellen. Klicken Sie im Visual Basic Editor, während ein Modul aktiv ist, im Menü  **Extras** auf **Verweise**, und wählen Sie dann unter  **Verfügbare Verweise** den Befehl **Solver** aus. Wenn **Solver** nicht unter **Verfügbare Verweise** angezeigt wird, klicken Sie auf **Durchsuchen**, und öffnen Sie dann  **Solver.xlam** im Unterordner **\Programme\Microsoft Office\Office14\Library\SOLVER**.


 **SolverLoad( _LoadArea_**, ** _Merge_ )**

 **LoadArea** Erforderlicher **Variant** -Wert. Ein Bezug im aktiven Arbeitsblatt auf einen Zellbereich, aus dem Sie eine vollständige Problemspezifikation laden möchten. Die erste Zelle in ** _LoadArea_** enthält eine Formel für das Feld **Zielzelle** im Dialogfeld **Solver-Parameter**; die zweite Zelle enthält eine Formel für das Feld **Veränderbare Zellen**; nachfolgende Zellen enthalten Nebenbedingungen in Form von logischen Formeln. Die letzte Zelle kann optional eine Liste von Solveroptionswerten enthalten. Weitere Informationen finden Sie unter **[SolverOptions-Funktion](270d5440-ac1e-2436-b632-5877ede0820e.md)**. Der durch das Argument ** _LoadArea_** dargestellte Bereich kann sich auf einem beliebigen Arbeitsblatt befinden. Sie müssen jedoch das Arbeitsblatt angeben, falls es sich nicht um das aktive Blatt handelt. Beispielsweise wird mit `SolverLoad("Sheet2!A1:A3")` ein Modell von Tabelle2 geladen, auch wenn dieses nicht das aktive Blatt ist.
 **Merge** Optionaler **Variant** -Wert. Ein logischer Wert, der entweder der Schaltfläche **Zusammenführen** oder der Schaltfläche **Ersetzen** im Dialogfeld entspricht, das nach dem Auswählen des Verweises **LoadArea** und Klicken auf **OK** angezeigt wird. Bei **True** wird die Auswahl von Variablenzellen und Nebenbedingungen aus LoadArea mit den aktuell definierten Variablen und Nebenbedingungen zusammengeführt. Wenn der Wert **False** oder nicht angegeben ist, werden die Angaben und Optionen zum aktuellen Modell gelöscht (gleichbedeutend mit einem Aufruf der Funktion **[SolverReset](5c8f99e7-9451-3e72-1d93-4fcd72fc3e71.md)** ), bevor die neuen Angaben geladen werden.

## Beispiel

In diesem Beispiel wird das zuvor berechnete und in Sheet1 gespeicherte Solvermodell geladen, eine der Nebenbedingungen geändert und das Modell erneut ausgewertet.


```
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38") 
SolverChange cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=200 
SolverSolve userFinish:=False
```

