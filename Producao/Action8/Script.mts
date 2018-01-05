'************************************************VERIFICAÇÃO DE ABA ATIVA************************************
'************************************************************************************************************

i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("ABA_COBERTURA")
obj.RefreshObject

If obj.Exist(0) Then
	If obj.GetROProperty("class") = "aba-ativa" Then
		If vezes_seguradora = 1 Then	
			DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
			Call ObtemHoraAtual(DataAtual, DataAtual2)
			Time_Questionario_Cobertura = tempo
			DataAtual = Empty
			DataAtual2 = Empty
		End If
		Exit Do
	End If
End If

If i = 20 Then
	Set obj = Browser("name:=Porto Print.*").Page("title:=Porto Print.*").WebElement("html id:=orcamentoAbaCobertura")
	If Clicar(obj) = False Then ExitAction(False) End If
End If

If i > 180 Then
	INCIDENTE = "Aba Cobertura Não Carregou no tempo limite" & i & "!"
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
	'*******************************************PORTO CONECTA****************************************************
	'************************************************************************************************************
	Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebRadioGroup("LINHA_CONECTA").Select "1"
	
	'*******************************************TIPO DE COBERTURA************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("TIPO_DE_COBERTURA")
	If CampoSetSel(obj,COBERTURA_COBERTURACASCO) = False Then ExitAction(False) End If
	'*********************************************FRANQUIA*******************************************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("FRANQUIA")
	If CampoSetSel(obj,COBERTURA_FRANQUIACASCO) = False Then ExitAction(False) End If
	
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DANOS_MATERIAIS")
	If CampoSetSel(obj,"50.000,00") = False Then ExitAction(False) End If
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("DANOS_CORPORAIS")
	If CampoSetSel(obj,"50.000,00") = False Then ExitAction(False) End If
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("MORTE_INVALIDEZ")
	If CampoSetSel(obj,"10.000,00") = False Then ExitAction(False) End If
End If

If vezes_seguradora = 2 Then
	'*********************************************RELATÓRIO - TAXA DE TRANSFERÊNCIA******************************
	'************************************************************************************************************
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Image("Relatório")
	If Clicar(obj) = False Then ExitAction(False) End If
	
	
		i = 0 
		Do
			Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("MSG_30S_TAXA_INTERNET")
			obj.RefreshObject
			If obj.Exist(0) Then
				If obj.GetROProperty("height") > 1 Then
					Exit Do
			    End If	
			End If
			
			If i > 80 Then
				INCIDENTE = "Modal de Taxa de Transferencia de dados(Relatório) não apareceu corretamente"
			End If
			
			wait(1)
			i = i + 1 
		Loop
		
		
		i = 0 
		Do
			Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("RELATORIO_TAXA")
			obj.RefreshObject
			If obj.Exist(0) Then
				If obj.GetROProperty("height") > 1 Then
					Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("VALOR_TAXA")
					TAXA_TRANSFERENCIA = obj.GetROProperty("innertext")
					Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("FECHAR_MODAL_TAXA_TRANSFERENCIA")
					If Clicar(obj) = False Then ExitAction(False) End If
					Exit Do
			    End If	
			End If
			
			If i > 80 Then
				INCIDENTE = "Modal de Taxa de Transferencia de dados(Relatório) não apareceu corretamente"
			End If
			
			wait(1)
			i = i + 1 
		Loop

End If

'*********************************************BOTÃO CONTINUAR************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Calcular")
If Clicar(obj) = False Then ExitAction(False) End If

DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
'*********************************************AGUARDE MODAL HISTÓRICO DE TRANSMISSÃO*************************
'************************************************************************************************************

i = 0 
Do
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("MODAL_AGUARDE")
	obj.RefreshObject
	If obj.Exist(0) Then
		If obj.GetROProperty("visible") = False Then
			Exit Do
	    End If	
	End If
	
	If i = 40 Then
		'*********************************************BOTÃO CONTINUAR************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Calcular")
		If Clicar(obj) = False Then ExitAction(False) End If
	End If
	
	If i > 180 Then
		INCIDENTE = "Tempo limite para aparecer o modal do Histórico de Transmissão do Orçamento"
		Exit Do
	End If
	
	wait 0, 200
	i = i + 1 
Loop

'*********************************************MODAL HISTÓRICO DE TRANSMISSÃO*********************************
'************************************************************************************************************

i = 0 
Do
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("HISTORICO_TRANSMISSAO")
	obj.RefreshObject
	If obj.GetROProperty("height") > 1 Then
		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("FECHAR_MODAL_TRANSMISSAO")
	    If Clicar(obj) = False Then ExitAction(False) End If
	    Exit Do
	End If

	If i > 10 Then
		INCIDENTE = "Modal Histórico de Transmissão não fechou corretamente"
		Exit Do
	End If
	
	wait(0.2)
	i = i + 1 
Loop


ExitAction(True)
