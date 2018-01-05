'************************************************VERIFICAÇÃO DE ABA ATIVA************************************
'************************************************************************************************************

i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("ABA_VEICULO")
obj.RefreshObject
If obj.Exist(0) Then
	If obj.GetROProperty("class") = "aba-ativa" Then
		If vezes_seguradora = 1 Then
			DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
			Call ObtemHoraAtual(DataAtual, DataAtual2)
			Time_Cadastro_Veiculo = tempo
			DataAtual = Empty
			DataAtual2 = Empty
		End If
		Exit Do
	End If
End If

If i > 20  Then
	obj.Click()
End If


If i > 180 Then
	INCIDENTE = "Aba Veículo Não Carregou no tempo limite" & i & "!"
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
	'************************************************FABRICACAO**************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("FABRICACAO")
	If CampoSetSel(obj,VEICULO_ANO) = False Then ExitAction(False) End If
	
	'************************************************MODELO******************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("MODELO")
	If CampoSetSel(obj,VEICULO_MODELO) = False Then ExitAction(False) End If
End If

wait(1)

'************************************************DESCRICAO***************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DESCRICAO")
If vezes_seguradora = 1 Then
	If CampoSetSel(obj,VEICULO_DESCRICAO1) = False Then ExitAction(False) End If
ElseIf vezes_seguradora = 2 Then 
	If CampoSetSel(obj,VEICULO_DESCRICAO2) = False Then ExitAction(False) End If
End If

i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DESCRICAO")
	If obj.GetROProperty("value") = "" Then	
		If vezes_seguradora = 1 Then
			If CampoSetSel(obj,VEICULO_DESCRICAO1) = False Then ExitAction(False) End If
		ElseIf vezes_seguradora = 2 Then 
			If CampoSetSel(obj,VEICULO_DESCRICAO2) = False Then ExitAction(False) End If
		End If
	Else
		Exit Do
	End If 
	If i > 30 Then
		Exit Do
	End If
wait(0.5)
i = i + 1
Loop
' **********AGUARDA MODAL DESCRICAO ABRIR, CARREGA VEICULO E CLICA EM CARREGAR******************************* 
i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("MODAL_VEICULO")
If obj.Exist(0) Then
	If obj.GetROProperty > 1 Then
		Set objbtn = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebRadioGroup("SELECT_VEICULO")
		If Clicar(objbtn) = False Then ExitAction(False) End If
		
		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("BTN_CARREGAR_VEIC")
		If Clicar(obj) = False Then ExitAction(False) End If
	
		Exit Do
	End If
End If

If i = 1 Then
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("BUSCAR_VEICULO")
	If Clicar(obj) = False Then ExitAction(False) End If
End If

If i > 10 Then
	TextoIncidente = "Modal Veiculo não apresentado"
	Exit Do
End If
wait(0.2)
i = i + 1
Loop

' **********AGUARDA MODAL DESCRICAO FECHAR**********************************************************************
i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("MODAL_VEICULO")
obj.RefreshObject
If obj.GetROProperty < 1 Then
wait(4)
	Exit Do
End If

If i > 100 Then
	TextoIncidente = "Modal Veiculo não fechou no tempo limite" & i
	Exit Do
End If
wait(0.2)
i = i + 1
Loop


'************************************************FABRICACAO**************************************************
'************************************************************************************************************
i = 0
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("FABRICACAO")
obj.RefreshObject
If obj.GetROProperty("value") = "" Then
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("FABRICACAO")
	If CampoSetSel(obj,VEICULO_ANO) = False Then ExitAction(False) End If
	Exit Do
End If

If i > 20 Then
	INCIDENTE = "Problemas ao preencher o Ano de Fabricação do Veículo!"
	Exit Do
End If
wait(0.2)
i = i + 1
Loop

'************************************************VALOR_BASE**************************************************
'************************************************************************************************************
'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("VALOR_BASE")
'If CampoSetSel(obj,"") = False Then ExitAction(False) End If


	'************************************************ALIENADO****************************************************
	'************************************************************************************************************
	wait(2)
If vezes_seguradora  = 1 or vezes_seguradora = 2 Then	
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("ALIENADO")
	If CampoSetSel(obj,"Não") = False Then ExitAction(False) End If
	wait(1)
	'************************************************KIT_GAS*****************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("KIT_GAS")
	If CampoSetSel(obj,"Não") = False Then ExitAction(False) End If
	wait(2)
	
	'*******************************PREENCHE NOVAMENTE O CAMPO "ALIENADO"****************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("ALIENADO")
	If CampoSetSel(obj,"Não") = False Then ExitAction(False) End If
	
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("KIT_GAS")
	If obj.GetROProperty("value") <> "Não" Then
		obj.Select("Não")
	End If

End If

If vezes_seguradora = 1 Then
'************************************************CAPTURA NUMERO DO ORÇAMENTO*********************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("NUMERO_ORCAMENTO")
NUMERO_ORCAMENTO = obj.GetROProperty("innertext")
Call GerenciadorGravaCampo("Numero_Orcamento", NUMERO_ORCAMENTO)

'************************************************BOTÃO AVANÇAR***********************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("BTN_CONTINUAR")
If Clicar(obj) = False Then ExitAction(False) End If

End If


If vezes_seguradora = 2 Then
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("ABA_COBERTURA")
	If Clicar(obj) = False Then ExitAction(False) End If
	wait(2)
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("ABA_COBERTURA")
	If Clicar(obj) = False Then ExitAction(False) End If
End If

DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)

ExitAction(True)


