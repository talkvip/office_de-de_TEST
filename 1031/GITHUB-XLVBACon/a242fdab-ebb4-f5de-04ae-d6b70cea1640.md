
# LegendEntry-Objekt

Stellt einen Legendeneintrag in der angegebenen Diagrammlegende. Das  **LegendEntry** -Objekt ist ein Mitglied der **[LegendEntries](98f5f860-7648-e3a6-f2af-985b383fed24.md)** -Auflistung, die alle **LegendEntry** -Objekte in der Legende enthält.

Jeder Legendeneintrag besteht aus zwei Teilen: der Text des Eintrags, der den Namen der Datenreihe dem Eintrag zugeordnet ist. und eine eintragsmarkierung, die mit der verknüpften Datenreihe oder Trendlinie im Diagramm den jeweiligen Legendeneintrag visuell verknüpft. Formateigenschaften der eingangsmarkierung und der zugeordneten Datenreihe oder Trendlinie sind im  **[LegendKey](ab90cb64-1f81-dfcb-7542-cba68964acba.md)** -Objekt enthalten.

Der Text eines Legendeneintrags kann nicht geändert werden.  **LegendEntry** -Objekte unterstützen Zeichenformatierung und können nicht gelöscht werden. Bei Legendeneinträgen wird Musterformatierung nicht unterstützt. Position und Größe von Einträgen sind feste Werte.


## Verwenden des LegendEntry-Objekts

Verwenden Sie  **LegendEntries** ( _Index_ ), wobei _Index_ die Indexnummer des Legendeneintrags ist, um ein einzelnes **LegendEntry** -Objekt zurückzugeben. Sie können Legendeneinträge nicht durch Angabe von Namen zurückgeben.

Die Indexnummer stellt die Position des Legendeneintrags in der Legende dar.  `LegendEntries(1)` wird am oberen Rand der Legende, steht oben in der Legende und `LegendEntries(LegendEntries.Count)` befindet sich unten. Im folgenden Beispiel wird den Schriftschnitt für den Text des Legendeneintrags am oberen Rand der Legende (dieser Schritt ist in der Regel die Legende für Datenreihe 1).




```
myChart.Legend.LegendEntries(1).Font.Italic = True
```


## Hinweise

Die Datenreihe und die Trendlinie, die zu einem bestimmten Legendeneintrag gehören, können auf direktem Weg nicht zurückgegeben werden.

Nach dem Legende, die Einträge gelöscht wurden, werden die einzige Möglichkeit zur Wiederherstellung entfernen und dann neu erstellen der Legende, die sie enthalten, indem Sie die  **[HasLegend](b4dbef39-9d83-2f6e-fe06-8ca38cceeeec.md)** -Eigenschaft für das Diagramm auf **false festgelegt** und dann auf **"true"** zurück.

