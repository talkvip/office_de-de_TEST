
# TextEffectFormat.FontItalic Property (Excel)

Returns  **msoTrue** if the font in the specified WordArt is italic. Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**.


## Syntax

 _Ausdruck_. **FontItalic**

 _Ausdruck_ A variable that represents a **TextEffectFormat** object.


## Remarks




||
|:-----|
|**MsoTriState** can be one of these **MsoTriState** constants.|
|**msoCTrue** Does not apply to this property.|
|**msoFalse** The specified WordArt is not italic.|
|**msoTriStateMixed** Does not apply to this property.|
|**msoTriStateToggle** Does not apply to this property.|
|**msoTrue** The specified WordArt is italic.|

## Example

This example sets the font to italic for the shape named "WordArt 4" in  `myDocument`.


```
Set myDocument = Worksheets(1) 
myDocument.Shapes("WordArt 4").TextEffect.FontItalic = msoTrue
```


## Siehe auch


#### Konzepte


[TextEffectFormat Object](7fe03721-6a45-569e-add4-fc8849c99535.md)
#### Weitere Ressourcen


[TextEffectFormat Object Members](http://msdn.microsoft.com/library/10d920d6-b96f-7afa-8e27-c22ba0926146%28Office.15%29.aspx)