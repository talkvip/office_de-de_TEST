
# WorksheetFunction.BesselY Method (Excel)

Returns the Bessel function, which is also called the Weber function or the Neumann function.


## Syntax

 _Ausdruck_. **BesselY**( ** _Arg1_**, ** _Arg2_** )

 _Ausdruck_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Erforderlich|**Variant**|The value at which to evaluate the function.|
| _Arg2_|Erforderlich|**Variant**|The order of the function. If n is not an integer, it is truncated.|

### Return Value

Double


## Remarks




- If x is nonnumeric, BesselY generates an error value.
    
- If n is nonnumeric, BesselY generates an error value.
    
- If n < 0, BesselY generates an error value.
    
- The n-th order Bessel function of the variable x is:
![](images/awfbsly1_ZA06051118.gif)


    

## Siehe auch


#### Konzepte


[WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Weitere Ressourcen


[WorksheetFunction Object Members](http://msdn.microsoft.com/library/6811ca87-4b53-0bff-88c9-30bf7497879a%28Office.15%29.aspx)