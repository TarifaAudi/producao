'************************************************VERIFICAÇÃO DE ABA ATIVA************************************
'************************************************************************************************************

i = 0
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DATA_NASCIMENTO")
obj.RefreshObject
If obj.Exist(0) Then
	If obj.GetROProperty("width") > 1 Then
		If vezes_seguradora = 1 Then
			DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
			Call ObtemHoraAtual(DataAtual, DataAtual2)
			Time_Apol_Cadastro = tempo
			DataAtual = Empty
			DataAtual2 = Empty
		End If
		Exit Do
	End If
End If

If i = 20  Then
	Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Apólice/Renovação").Click @@ hightlight id_;_Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Apólice/Renovação")_;_script infofile_;_ZIP::ssf1.xml_;_
End If

If i > 180 Then
	INCIDENTE = "Aba Cadastro Não Carregou no tempo limite" & i & "!"
	GerenciadorGravaCampo "Incidente", INCIDENTE
	GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
	ErroAction = True
	ExitAction(False)
	Exit Do
End If

wait(1)
i = i + 1
Loop


If vezes_seguradora = 1 Then
	Set SegPorto = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SEGURADORA_PORTO")
	If SegPorto.GetROProperty("checked") = 0 Then
		If Clicar(SegPorto) = False Then ExitAction(False)
	End If
	
	Set SegItau = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SEGURADORA_ITAU")
	If SegItau.GetROProperty("checked") = 0 Then
		If Clicar(SegItau) = False Then ExitAction(False)
	End If
	
	Set SegAzul = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SEGURADORA_AZUL")
	If SegAzul.GetROProperty("checked") = 0 Then
		If Clicar(SegAzul) = False Then ExitAction(False)
	End If
	
ElseIf vezes_seguradora = 2 Then 
	Set SegItau = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SEGURADORA_ITAU")
	If SegItau.GetROProperty("checked") = 1 Then
		If Clicar(SegItau) = False Then ExitAction(False)
	End If
End If

If vezes_seguradora = 1 Then
	'************************************************CPF*********************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CPF")
	If CampoSetSel(obj,CADASTRO_CPF) = False Then ExitAction(False) End If
	wait(3)
	
	'************************************************NOME********************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("NOME")
	If CampoSetSel(obj,CADASTRO_NOME) = False Then ExitAction(False) End If
	'********************************************DATA_NASCIMENTO*************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DATA_NASCIMENTO")
	
	dia = mid(CADASTRO_DATANASCIMENTO,1,2)
	mes = mid(CADASTRO_DATANASCIMENTO,3,2)
	ano = mid(CADASTRO_DATANASCIMENTO,5,4)
	
	CADASTRO_DATANASCIMENTO = dia & "/" & mes & "/" & ano
	
	
	If CampoSetSel(obj,CADASTRO_DATANASCIMENTO) = False Then ExitAction(False) End If
	'************************************************CEP_NUMERO**************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CEP_NUMERO")
	If CampoSetSel(obj,CADASTRO_CEP) = False Then ExitAction(False) End If
	'************************************************CEP_COMPLEMENTO*********************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CEP_COMPLEMENTO")
	If CampoSetSel(obj,CADASTRO_CEP_COMPLEMENTO) = False Then ExitAction(False) End If

End If
'***************************************BOTÃO AVANÇAR********************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("BTN_CONTINUAR")
If Clicar(obj) = False Then ExitAction(False) End If

If vezes_seguradora = 1 Then
	DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
End If
ExitAction(True)
