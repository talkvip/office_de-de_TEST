
# DefaultWebOptions.AllowPNG Property (Excel)

 **True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a Web page. **False** if PNG is not allowed as an output format. The default value is **False**. Read/write **Boolean**.


## Syntax

 _Ausdruck_. **AllowPNG**

 _Ausdruck_ A variable that represents a **DefaultWebOptions** object.


## Remarks

If you save images in the PNG format as opposed to any other file format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time, assuming that the Web browsers you are targeting support the PNG format.


## Example

Alternatively, PNG can be enabled as the global default for the application for newly created documents.


```
Application.DefaultWebOptions.AllowPNG = True
```


## Siehe auch


#### Konzepte


[DefaultWebOptions Object](5bd1d870-e8d9-cac1-d7a7-3aeaf7c4c3cd.md)
#### Weitere Ressourcen


[DefaultWebOptions Object Members](http://msdn.microsoft.com/library/52db1398-01d8-eba5-772f-2923fdc89f5b%28Office.15%29.aspx)