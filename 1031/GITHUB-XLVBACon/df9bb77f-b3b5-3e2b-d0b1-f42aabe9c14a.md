
# WorksheetFunction.CritBinom Method (Excel)

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.


## Syntax

 _Ausdruck_. **CritBinom**( ** _Arg1_**, ** _Arg2_**, ** _Arg3_** )

 _Ausdruck_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Erforderlich|**Double**|The number of Bernoulli trials.|
| _Arg2_|Erforderlich|**Double**|The probability of a success on each trial.|
| _Arg3_|Erforderlich|**Double**|The criterion value.|

### Return Value

Double


## Remarks

 Use this function for quality assurance applications. For example, use CritBinom to determine the greatest number of defective parts that are allowed to come off an assembly line run without rejecting the entire lot.


- If any argument is nonnumeric, CritBinom generates an error.
    
- If trials is not an integer, it is truncated.
    
- If trials < 0, CritBinom generates an error.
    
- If probability_s is < 0 or probability_s > 1, CritBinom generates an error.
    
- If alpha < 0 or alpha > 1, CritBinom generates an error.
    

## Siehe auch


#### Konzepte


[WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Weitere Ressourcen


[WorksheetFunction Object Members](http://msdn.microsoft.com/library/6811ca87-4b53-0bff-88c9-30bf7497879a%28Office.15%29.aspx)