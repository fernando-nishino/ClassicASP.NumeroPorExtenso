<%
'Language: PT-BR
'https://github.com/fernando-nishino/ClassicASP.NumeroPorExtenso

Response.CharSet = "utf-8"
Class NumeroPorExtenso
	Private npeNumero
	Private npeMasculino
	Private npeParteNumero()
	
	Public Property Let Numero(valor)
		npeNumero = valor
	End Property
	Public Property Let Masculino(valor)
		npeMasculino = valor
	End Property
	
	Private Sub Class_Initialize()		
		npeMasculino = True
	End Sub
	
	Private Function NomeNumero(ByVal numero, ByVal variacao, ByVal plural)
		Select Case (numero)
			Case 0  NomeNumero = "zero"
			Case 1  If npeMasculino Then NomeNumero = "um" 		Else NomeNumero = "uma"
			Case 2  If npeMasculino Then NomeNumero = "dois" 	Else NomeNumero = "duas"
			Case 3  NomeNumero = "três"
			Case 4  NomeNumero = "quatro"
			Case 5  NomeNumero = "cinco"
			Case 6  NomeNumero = "seis"
			Case 7  NomeNumero = "sete"
			Case 8  NomeNumero = "oito"
			Case 9  NomeNumero = "nove"
			Case 10 NomeNumero = "dez"
			Case 11 NomeNumero = "onze"
			Case 12 NomeNumero = "doze"
			Case 13 NomeNumero = "treze"
			Case 14 NomeNumero = "quatorze" 'ou "catorze"
			Case 15 NomeNumero = "quize"
			Case 16 NomeNumero = "dezesseis"
			Case 17 NomeNumero = "dezessete"
			Case 18 NomeNumero = "dezoito"
			Case 19 NomeNumero = "dezenove"
			Case 20 NomeNumero = "vinte"
			Case 30 NomeNumero = "trinta"
			Case 40 NomeNumero = "quarenta"
			Case 50 NomeNumero = "cinquenta"
			Case 60 NomeNumero = "sessenta"
			Case 70 NomeNumero = "setenta"
			Case 80 NomeNumero = "oitenta"
			Case 90 NomeNumero = "noventa"
			Case 100 If variacao Then NomeNumero = "cento" Else NomeNumero = "cem"
			Case 200 If npeMasculino 	Then NomeNumero = "duzentos" 		Else NomeNumero = "duzentas"
			Case 300 If npeMasculino 	Then NomeNumero = "trezentos" 		Else NomeNumero = "trezentas"
			Case 400 If npeMasculino 	Then NomeNumero = "quatrocentos" 	Else NomeNumero = "quatrocentas"
			Case 500 If npeMasculino 	Then NomeNumero = "quinhentos" 		Else NomeNumero = "quinhentas"
			Case 600 If npeMasculino 	Then NomeNumero = "seiscentos" 		Else NomeNumero = "seiscentas"
			Case 700 If npeMasculino 	Then NomeNumero = "setecentos" 		Else NomeNumero = "setecentas"
			Case 800 If npeMasculino 	Then NomeNumero = "oitocentos" 		Else NomeNumero = "oitocentas"
			Case 900 If npeMasculino 	Then NomeNumero = "novecentos" 		Else NomeNumero = "novecentas"
			Case 1000 NomeNumero = "mil"
			Case 1000000 	If plural Then NomeNumero = "milhões" Else NomeNumero = "milhão"
			Case 1000000000 If plural Then NomeNumero = "bilhões" Else NomeNumero = "bilhão"
		End Select
	End Function

	Private Sub Separar
		Dim numero
		ReDim npeParteNumero(0)
		numero = npeNumero
		While numero >= 1
			npeParteNumero(UBound(npeParteNumero)) = numero Mod 1000
			numero = numero \ 1000
			ReDim Preserve npeParteNumero(UBound(npeParteNumero) + 1)
		Wend
	End Sub
		
	Private Function Regra()
		Dim parte, i, milhares, resultado
		For i = UBound(npeParteNumero) - 1 To 0 Step - 1
			parte = npeParteNumero(i)
			If parte > 0 Then
				milhares = ""
				If i > 0 Then
					If parte > 1 Then
						milhares = " " & NomeNumero(1000 ^ i, False, True)
					Else
						milhares = " " & NomeNumero(1000 ^ i, False, False)
					End If
				End If
				
				If i = 0 And UBound(npeParteNumero) > 1 Then
					If parte Mod 100 = 0 Or parte < 100 Then
						resultado = resultado & " e "
					End If
				End If
				
				If parte Mod 100 = 0 Then
					resultado = resultado & NomeNumero(parte, False, False) & milhares & " "
				ElseIf parte < 20 Then
					resultado = resultado & NomeNumero(parte, True, False) & milhares & " "
				ElseIf parte > 0 Then
					If parte > 100 Then
						resultado = resultado & NomeNumero((parte \ 100) * 100, True, False) & " "
						parte = parte Mod 100
						If parte > 0 Then resultado = resultado & " e "
					End If
					If parte >= 20 Then
						resultado = resultado & NomeNumero((parte \ 10) * 10, True, False) & " "
						If parte Mod 10 > 0 Then
							resultado = resultado & " e " & NomeNumero(parte Mod 10, True, False)
						End If
					End If
					resultado = resultado & milhares & " "
				End If
			End If
		Next
		Regra = resultado
	End Function
	
	Private Function Gerar()
		Separar
		Gerar = Regra
	End Function
	
	'genero feminino pode ser ("f", "feminino", "fem") ou false. Qualquer outro valor vai ser masculino
	Public Function Retornar(ByVal numero, ByVal genero)
		If numero > 0 Then
			npeNumero = numero
			If LCase(Left(genero, 1)) = "f" Or Not genero Then
				npeMasculino = False
			End If
			Retornar = Gerar
		Else
			Retornar = NomeNumero(0, False, False)
		End If
	End Function
End Class


Set numeroExtenso = New NumeroPorExtenso
Response.Write numeroExtenso.Retornar(1234567890, True) & "<br>"
Response.Write numeroExtenso.Retornar(21030002, True) & "<br>"
Response.Write numeroExtenso.Retornar(1065, True) & "<br>"
Response.Write numeroExtenso.Retornar(75600, True) & "<br>"
Set numExt = Nothing


%>
