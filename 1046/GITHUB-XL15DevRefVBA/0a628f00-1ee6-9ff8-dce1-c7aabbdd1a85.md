
# Método Chart.SeriesCollection (Excel)

Retorna um objeto que representa uma única série (um objeto  **[Series](c7d34b32-8172-f7a0-0a17-f01d44246b64.md)** ) ou uma coleção de todas as séries (uma coleção **[SeriesCollection](93aa1f0b-4939-8c60-a444-2f791e8ce144.md)** ) no gráfico ou grupo de gráficos.


## Sintaxe

 _expressão_. **SeriesCollection**( ** _Index_** )

 _expressão_ Uma variável que representa um objeto **Chart**.


### Parâmetros



|**Nome**|**Obrigatório/opcional**|**Tipo de dados**|**Descrição**|
|:-----|:-----|:-----|:-----|
| _Index_|Opcional|**Variant**|O nome ou número da série.|

### Valor retornado

Objeto


## Exemplo

Este exemplo ativa os rótulos de dados para a série um em Chart1.


```
Charts("Chart1").SeriesCollection(1).HasDataLabels = True
```


## Ver também


#### Conceitos


[Objeto Chart](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)
#### Outros recursos


[Membros do Objeto Chart](a3f8ac44-02d6-6f3f-b5e0-23f4bd5d6baf.md)