
# Função FormatDateTime



 **Descrição**
Retorna uma expressão formatada como uma data ou hora.
 **Sintaxe**
 **FormatDateTime(** _Date_ [ **,** _NamedFormat_ ] **)**
A sintaxe da função  **FormatDateTime** tem as três partes abaixo:


|**Parte**|**Descrição**|
|:-----|:-----|
| _Data_|Obrigatório. Expressão de data a ser formatada.|
| _NamedFormat_|Opcional. Valor numérico que indica o formato de data/hora usado. Se for omitido,  **vbGeneralDate** será usado.|
 **Configurações**
O argumento  _NamedFormat_ apresenta as seguintes configurações:


|**Constante**|**Valor**|**Descrição**|
|:-----|:-----|:-----|
|**vbGeneralDate**|0|Exibe uma data e/ou hora. Se houver uma parte de data, exibi-la como uma data abreviada. Se houver uma parte de hora, exibi-la como uma hora por extenso. Se estiver presente, ambas as partes são exibidas.|
|**vbLongDate**|1|Exibe uma data usando o formato de data por extenso especificado nas configurações regionais do seu computador.|
|**vbShortDate**|2|Exibe uma data usando o formato de data abreviado especificado nas configurações regionais do seu computador.|
|**vbLongTime**|3|Exibe uma hora usando o formato de hora especificado nas configurações regionais do seu computador.|
|**vbShortTime**|4|Exibe uma hora usando o formato de 24 horas (hh:mm).|
