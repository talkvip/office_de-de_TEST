
# ODBCConnection.SourceDataFile Property (Excel)

Returns or sets a  **String** indicating the source data file for an ODBC connection. Read/write.


## Syntax

 _Ausdruck_. **SourceDataFile**

 _Ausdruck_ A variable that represents an **ODBCConnection** object.


## Remarks

For file-based data sources (for example, Access) the  **SourceDataFile** property contains a fully qualified path to the source data file. It is null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to null if the **[Connection](2fcd1043-b088-cfde-9853-4a20da20be26.md)** property is changed programmatically.


## Siehe auch


#### Konzepte


[ODBCConnection Object](b880ebec-15a4-5a3d-ef02-db73106db9c9.md)
#### Weitere Ressourcen


[ODBCConnection Object Members](http://msdn.microsoft.com/library/d13b91f3-a89f-7dd7-7a98-f1d952f3b047%28Office.15%29.aspx)