
# Workbook.BeforeXmlExport Event (Excel)

Occurs before Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

 _Ausdruck_. **BeforeXmlExport**( ** _Map_**, ** _Url_**, ** _Cancel_** )

 _Ausdruck_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Map_|Erforderlich|**[XmlMap](39b0823f-0068-d8df-e4e1-ca62b55d58f5.md)**|The XML map that will be used to save or export data.|
| _Url_|Erforderlich|**String**|The location where you want to export the resulting XML file.|
| _Cancel_|Erforderlich|**Boolean**|Set to  **True** to cancel the save or export operation|

### Return Value

Nothing


## Remarks

This event will not occur when you are saving to the XML Spreadsheet file format.

