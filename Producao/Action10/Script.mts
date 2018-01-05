
'************************************************REALIZA LEITURA DOS DADOS GA********************************
'************************************************************************************************************
Call LerDados()

ErroAction = False
vezes_seguradora = 1
Do
	If vezes_seguradora > 2 And ErroAction = False Then
		vezes_seguradora = Empty
		
		If ErroAction = False Then
			RunAction "Finaliza_Testes_Time", oneIteration
		Else
			ExitAction(True)
		End If

		DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
		Call GerenciadorGravaCampo("Termino_Execucao", DataAtual)
		Call GerenciadorFinalizaTeste()
		
		INCIDENTE = Empty
		Exit Do
	Else
		If ErroAction = True Then
			ExitAction(True)			
		Else
			Call FluxoAutomation()
		End If
	End If
	vezes_seguradora = vezes_seguradora + 1
Loop

Function FluxoAutomation()
	'************************************************INICIAR NAVEGAÇÃO*******************************************
	'************************************************************************************************************
	'RunAction("IniciarNavegacao")
	'************************************************LOGIN*******************************************************
	'************************************************************************************************************
	'RunAction("Login")
	If vezes_seguradora = 1 Then
		InicioTeste = now()
	'***********************************************PORTAL COL***************************************************
	'************************************************************************************************************
		If ErroAction = False Then
			RunAction "PortalCol", oneIteration
		End If
	End If
	'************************************************CADASTRO****************************************************	
	If ErroAction = False Then
		RunAction "Cadastro" , oneIteration
	End If
	'************************************************VEÍCULO*****************************************************
	If ErroAction = False Then
		RunAction "Veiculo", oneIteration
	End If
	'************************************************QUESTIONÁRIO************************************************
	If vezes_seguradora = 1 Then
		If ErroAction = False Then
			RunAction "Questionario", oneIteration
		End If
	End If
	'************************************************COBERTURA***************************************************
	If ErroAction = False Then
		RunAction "Cobertura", oneIteration
	End If
	'************************************************CÁLCULO*****************************************************
	If ErroAction = False Then
		RunAction "Calculo", oneIteration
	End If
End Function

ExitAction(True)
