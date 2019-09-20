# ClassicASP.NumeroPorExtenso

Escrever um número por extenso (PT-BR)

_Ex.: 1253492 = um milhão duzentos e cinquenta e três mil quatrocentos e noventa e dois_

Também é possível definir masculino ou feminino

_Ex.: 1253492 = um milhão duzentos e cinquenta e três mil quatrocentos e noventa e duas_

.

**Sintaxe:**

obj.Retornar (int, bool|string)

int = o número

bool ou string = "f" ou "m" ou True(masc) ou False(fem)

.

Set obj = New NumeroPorExtenso

Response.Write obj.Retornar(1234567891, False)
