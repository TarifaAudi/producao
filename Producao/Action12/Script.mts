	
'****************************************CLICA EM SAIR*********************************************************************

'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebButton("Sair_Encerra_Sessao")
'If Clicar(obj) = False Then ExitAction(False) End If

'Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebButton("Sair_Encerra_Sessao").Click
'****************************************FECHA ABA SESSAO******************************************************************
Set obj = Browser("name:=Porto Print.*")
If obj.Exist(1) Then
	obj.Close
End If

If PROJETO = "30_Auditeste" Then
	VelocidadeSpeedy = "30Mbp/s"
ElseIf PROJETO = "02_Porto" Then
	VelocidadeSpeedy = "02Mbp/s"
End If

'SystemUtil.CloseProcessByName("chrome.exe")

'****************************************VERIFICA SE PRECISA ENVIAR EMAIL**************************************************
'***************************************************************************************************************************
If Time_Portal_Apol_Renov > 07 Then
	TEMPO1 = "ACESSO DO COL PARA TELA DE APOLICE/RENOVAÇÃO | SLA ESPERADO:[7s] - SLA OBTIDO:[" & Time_Portal_Apol_Renov & "]" + vbLf + ""
Else
	TEMPO1 = ""
End If

If Time_Apol_Cadastro > 04 Then
	TEMPO2 = "ACESSO DO APOLICE/RENOVAÇÃO PARA TELA DE CADASTRO | SLA ESPERADO:[4s] - SLA OBTIDO:[" & Time_Apol_Cadastro & "]" + vbLf + ""
Else
	TEMPO2 = ""
End If

If Time_Cadastro_Veiculo > 05 Then
	TEMPO3 = "ACESSO DO CADASTRO PARA TELA DE VEÍCULO | SLA ESPERADO:[5s] - SLA OBTIDO:[" & Time_Cadastro_Veiculo & "]" + vbLf + ""
Else
	TEMPO3 = ""
End If

If Time_Veiculo_Questionario > 04 Then
	TEMPO4 = "ACESSO DO VEÍCULO PARA TELA DE QUESTIONÁRIO | SLA ESPERADO:[4s]- SLA OBTIDO:[" & Time_Veiculo_Questionario & "]" + vbLf + ""
Else
	TEMPO4 = ""
End If

If Time_Questionario_Cobertura > 05 Then
	TEMPO5 = "ACESSO DO QUESTIONÁRIO PARA TELA DE COBERTURA | SLA ESPERADO:[5s]- SLA OBTIDO:[" & Time_Questionario_Cobertura & "]" + vbLf + ""
Else
	TEMPO5 = ""	
End If

If Time_Cobertura_Calculo_Tres_Seg > 13 Then
	TEMPO6 = "TELA DE CALCULO COM 3 SEGURADORAS | SLA ESPERADO:[13s]- SLA OBTIDO:[" & Time_Cobertura_Calculo_Tres_Seg & "]" + vbLf + ""
Else
	TEMPO6 = ""	
End If

If Time_Cobertura_Calculo_Duas_Seg > 15 Then
	TEMPO7 = "TELA DE CALCULO COM 2 SEGURADORAS | SLA ESPERADO:[15s] - SLA OBTIDO:[" & Time_Cobertura_Calculo_Duas_Seg & "]" + vbLf + ""
Else
	TEMPO7 = ""		
End If

If TEMPO1 <> "" or TEMPO2 <> "" or TEMPO3 <> "" or TEMPO4 <> "" or TEMPO5 <> "" or TEMPO6 <> "" or TEMPO7 <> "" Then
		
	'****************************************ABRE O LINK DO WEBMAIL AUDITESTE***************************************************
	'***************************************************************************************************************************
	SystemUtil.Run "https://email.uolhost.com.br/auditeste.com.br/"
	wait 2
	'Browser("Porto Print - Porto Seguro").Navigate("https://email.uolhost.com.br/auditeste.com.br/")
	'SystemUtil.Run "https://email.uolhost.com.br/auditeste.com.br/"
	
	'****************************************PREENCHE USUÁRIO DO EMAIL**********************************************************
	'***************************************************************************************************************************
	Set obj = Browser("Webmail").Page("Webmail").WebElement("Seja bem vindo ao UOLHOST-")
	If Clicar(obj) = False Then 
		ExitAction(False) 
	Else
		wait 1
		Set obj = Browser("Webmail").Page("Webmail").WebEdit("Digite seu e-mail neste")
		If CampoSetSel(obj,"monitoracao") = False Then ExitAction(False) End If
	End If
	
	'****************************************PREENCHE A SENHA DO EMAIL**********************************************************
	'***************************************************************************************************************************
	Set obj = Browser("Webmail").Page("Webmail").WebEdit("Digite sua senha neste")
	If CampoSetSel(obj,"audi2017") = False Then ExitAction(False) End If
	
	'****************************************CLICA EM ENTRAR NO EMAIL***********************************************************
	'***************************************************************************************************************************
	Set obj = Browser("Webmail").Page("Webmail").WebButton("Entrar")
	If Clicar(obj) = False Then ExitAction(False) End If
	
	
	'************************************************VERIFICAÇÃO DE EMAIL ABERTO*********************************
	'************************************************************************************************************
	
	i = 0 
	Do
	Set obj = Browser("Porto Print - Porto Seguro").Page("Entrada - monitoracao@auditest").WebElement("ESCREVER_EMAIL")
	obj.RefreshObject
	
	If obj.GetROProperty("width") > 1 Then
		
		Exit Do
	End If
	
	If i > 180 Then
		INCIDENTE = "Webmail não carregou no tempo:" & i & "s!"
		GerenciadorGravaCampo "Incidente", INCIDENTE
		GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
		ErroAction = True
		ExitAction(False)
		Exit Do
	End If
	
	wait(1)
	i = i + 1
	Loop
	
	
	'************************************************TEXTOS PARA CORPO EMAIL*******************************************
	'******************************************************************************************************************
		FraseEmail = TEMPO1 & TEMPO2 & TEMPO3 & TEMPO4 & TEMPO5 & TEMPO6 & TEMPO7
		
		Texto_Email = "*******************SLA EXCEDIDO*******************" + vbLf + vbLf + "Tempo de SLA Excedido nas seguintes transições de abas:" + vbLf + "" & FraseEmail & "Tipo de conexão: Speedy - " & VelocidadeSpeedy  + vbLf + "" + vbLf + "Número do Documento: " & NUMERO_ORCAMENTO + vbLf + "Número do Orçamento: " & NUMERO_DOCUMENTO + vbLf + "" + vbLf + "Data/Hora - Inicio do Teste : " & InicioTeste & + vbLf + "Hora - Alerta: " & now() & + vbLf + "" + vbLf + vbLf + "Taxa de Transferencia :  " & TAXA_TRANSFERENCIA & + vbLf + "" + vbLf + "Obs: Quando o teste é interrompido antes da conclusão do cadastro, o número do orçamento não é gerado."
		
		Call EnviaEmail()
End If

			
	Function EnviaEmail()
		'************************************************USUÁRIOS PARA EMAIL***********************************************
		'******************************************************************************************************************
		Usuarios_Email_Porto = "jhon.callisaya@auditeste.com.br"
		Usuarios_Email_Auditeste = "jhon.callisaya@auditeste.com.br"'"jhon.callisaya@auditeste.com.br, silvia.dangelo@auditeste.com.br, allan.brito@auditeste.com.br"
					
		'************************************************BOTAO ESCREVER EMAIL**********************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Entrada - monitoracao@auditest").WebElement("ESCREVER_EMAIL")
		If Clicar(obj) = False Then ExitAction(False) End If
		
		'************************************************PARA:*************************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebEdit("NOVO_EMAIL_PARA")
		If CampoSetSel(obj,Usuarios_Email_Porto) = False Then ExitAction(False) End If
		
		'************************************************BOTAO COM CÓPIA***************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebElement("NOVO_EMAIL_COPIA_BTN")
		If Clicar(obj) = False Then ExitAction(False) End If
		wait 2
		
		'************************************************TEXTO COM CÓPIA***************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebEdit("NOVO_EMAIL_COPIA_TEXTO")
		If CampoSetSel(obj,Usuarios_Email_Auditeste) = False Then ExitAction(False) End If
		
		'************************************************ASSUNTO EMAIL*****************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebEdit("NOVO_EMAIL_ASSUNTO")
		If CampoSetSel(obj,"TEMPO DE SLA EXCEDIDO") = False Then ExitAction(False) End If
		
		'************************************************TEXTO CORPO EMAIL*************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebEdit("AREA_TEXTO_MENSAGEM")
		If CampoSetSel(obj,Texto_Email) = False Then ExitAction(False) End If
		
		'************************************************ENVIAR EMAIL******************************************************
		'******************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Novo e-mail - monitoracao@audi").WebElement("ENVIAR_EMAIL")
		If Clicar(obj) = False Then ExitAction(False) End If
		wait 3
		
		'FECHAR PAGINA WEBMAIL
		Set obj = Browser("name:=Novo e-mail.*")
		If obj.Exist(0) Then
			obj.Close
		End If
		Set obj = Browser("title:=Entrada.*")
		If obj.Exist(0) Then
			obj.Close
		End If
		Set obj = Browser("title:=Webmail.*")
		If obj.Exist(0) Then
			obj.Close
		End If
		Set obj = Browser("name:=Webmail.*")
		If obj.Exist(0) Then
			obj.Close
		End If		
		Call GerenciadorFinalizaTeste()
		
	End Function

'****************************************EXECUTA PROCESSO DE TIME (11 MINUTOS)**********************************************
'***************************************************************************************************************************

'SystemUtil.Run "D:\Producao\monitoramento"
'wait 660 '11 minutos, tempo de espera para o próximo teste 660 segundos
'SystemUtil.CloseProcessByName("chrome.exe")



'AGUARDAR OS 11 MINUTOS PARA EXECUTAR O PROXIMO TESTE
Call AtualizarPaginaAguardar()

SystemUltil.CloseProcessByName("D:\Producao\monitoramento")
'RunAction "GA"

Public Function AtualizarPaginaAguardar()
Dim iloop
iloop = 0
Do 

	If iloop > 9 Then
		Exit Do
	End If
	
	If Browser("title:=Porto Seguro.*").Exist Then
		tecla.SendKeys("{F5}")
	End If
	wait 60
	
	iloop = iloop + 1
Loop




End Function

