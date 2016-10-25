
# Função Date



Retorna uma  **Variant** ( **Date** ) com a data atual do sistema.
 **Syntax**
 **Date**
 **Remarks**
Para definir a data do sistema, use a instrução  **Date**.
 **Date** e, se o calendário for o Gregoriano, o comportamento **Date$** não será alterado pela configuração da propriedade **Calendar**. Se o calendário for o Hijri, **Date$** retornará uma cadeia de 10 caracteres no formato _mm-dd-aaaa_, em que _mm_ (01-12), _dd_ (01-30) e _aaaa_ (1400-1523) são o mês, o dia e o ano Hijri. O intervalo gregoriano equivalente é 1º de janeiro de 1980 a 31 de dezembro de 2099.

## Exemplo

Este exemplo usa a função  **Date** para retornar a data atual do sistema.


```
Dim MyDate
MyDate = Date    ' MyDate contains the current system date.


```

