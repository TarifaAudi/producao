'************************************************VERIFICAÇÃO DE ABA ATIVA************************************
'************************************************************************************************************

i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("OPCAO_DE_CALCULO")
obj.RefreshObject

If obj.GetROProperty("width") > 1 Then
	DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
	Call ObtemHoraAtual(DataAtual, DataAtual2)
	If vezes_seguradora = 1 Then
		Time_Cobertura_Calculo_Tres_Seg = tempo
	ElseIf vezes_seguradora = 2 Then 
		Time_Cobertura_Calculo_Duas_Seg = tempo
	End If
	
	DataAtual = Empty
	DataAtual2 = Empty
	Exit Do
End If


If i > 60 Then
	INCIDENTE = "Aba Calculo Não Carregou!"
	GerenciadorGravaCampo "Incidente", INCIDENTE
	GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
	ErroAction = True
	ExitAction(False)
	Exit Do
End If

wait(1)
i = i + 1
Loop


'******************************* GRAVA VALOR DA DESCRIÇÃO DO VEÍCULO NA VARIAVEL*****************************************************
'***************************************************************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("VEICULO_DESCRICAO_CALCULO")
		If obj.Exist(0) Then 
			Calculo_Veiculo_Descricao = ""
			Calculo_Veiculo_Descricao = obj.GetROProperty("innertext")
'		Else 
'			Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("ORCAMENTO_EM_PROCESSAMENTO")
'			If obj.Exist(0) Then
'				Orcamento_Processamento = obj.GetROProperty("text")
'				TextoIncidente = Orcamento_Processamento     
'				Call GerenciadorGravaCampo("Aceitacao",TextoIncidente )
'				If ErrorTest (obj, TextoIncidente) = False Then ExitAction(False)
'				ExitAction (False)
'			End If
		End If

'******************************* GRAVA VALOR DA CLASSE DE LOCALIZAÇÃO DO VEÍCULO NA VARIAVEL************************************
'***************************************************************************************************************************************************************
		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("CLASSE_LOCALIZACAO")
		Calculo_ClasseLocalizacao_Descricao = ""
		Calculo_ClasseLocalizacao_Descricao = obj.GetROProperty("innertext")

'******************************* CAPTURA VALORES DO PREMIO******************************************************************
'***************************************************************************************************************************************************************

		Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebTable("OPCAO_DE_CALCULO")
		linhatotal = obj.RowCount()

		linha = 1

		Do	
			If linha > linhatotal Then
				Exit Do
			Else 
				 If  obj.GetCellData(linha, 1) = "Prêmio Auto" Then
					If vezes_seguradora = 1 Then
						premio_porto_tres = obj.GetCellData(linha, 2)
					 	premio_itau_tres = obj.GetCellData(linha, 3)
					 	premio_azul_tres = obj.GetCellData(linha, 4)
					 	
					 	Calculo_Aceitacao_Tres_Seg = "3 SEGURADORAS - Padrão  - Opção de Cálculo:"&"(" &Calculo_Veiculo_Descricao & ")"&"Classe de Localização:" & "(" & Calculo_ClasseLocalizacao_Descricao & ")" & "Valor do Prêmio 3 Seguradoras " & "(" & "PORTO:[" & premio_porto_tres & "]" & "ITAÚ:[" & premio_itau_tres & "]" & "AZUL:[" & premio_azul_tres & "]" & ")"
					 	
					ElseIf vezes_seguradora = 2 Then 
						premio_porto_duas = obj.GetCellData(linha, 2)
					 	premio_itau_duas = obj.GetCellData(linha, 3)
					 	
					 	Calculo_Aceitacao_Duas_Seg = " 2 SEGURADORAS - Padrão  - Opção de Cálculo:"&"(" &Calculo_Veiculo_Descricao & ")"&"Classe de Localização:" & "(" & Calculo_ClasseLocalizacao_Descricao & ")" & "Valor do Prêmio" & "(" & "PORTO:[" & premio_porto_tres & "]" & "ITAÚ:[" & premio_itau_tres & "]" & ")"
					End If
				  End If
			 End If
		linha = linha + 1
		Loop 

'******************************* CAPTURA NUMERO DO DOCUMENTO******************************************************************************
'***************************************************************************************************************************************************************
If vezes_seguradora = 1 Then
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("NUMERO_COTACAO")
	NUMERO_DOCUMENTO = obj.GetROProperty("innertext")
	Call GerenciadorGravaCampo("Numero_Documento", NUMERO_DOCUMENTO)
End If	

'******************************* GRAVA ACEITAÇÃO DO VALOR PRÊMIO*****************************************************************************
'********************************************************************************************************************************************
Call GerenciadorGravaCampo("Aceitacao", Calculo_Aceitacao_Tres_Seg & Calculo_Aceitacao_Duas_Seg)


'******************************* GRAVA TEMPOS DAS ABAS***************************************************************************************
'********************************************************************************************************************************************

Call GerenciadorGravaCampo("Primeiro_Acesso", "00:00:" & Time_Portal_Apol_Renov)
Call GerenciadorGravaCampo("Novo_Orcamento", "00:00:" & Time_Apol_Cadastro)

Call GerenciadorGravaCampo("Cadastro_Veiculo", "00:00:" & Time_Cadastro_Veiculo)
Call GerenciadorGravaCampo("Veiculo_Perfil", "00:00:" & Time_Veiculo_Questionario)
Call GerenciadorGravaCampo("Perfil_Cobertura", "00:00:" & Time_Questionario_Cobertura)
Call GerenciadorGravaCampo("Tres_Seguradoras", "00:00:" & Time_Cobertura_Calculo_Tres_Seg)
If vezes_seguradora = 1 Then
	Call printTela(objPrint, "Evidencia_Tres_Seguradoras")
End If
If vezes_seguradora = 2 Then
	Call GerenciadorGravaCampo("Duas_Seguradoras",	"00:00:" & Time_Cobertura_Calculo_Duas_Seg)
	Call printTela(objPrint, "Evidencia_Duas_Seguradoras")
	Call GerenciadorGravaCampo("Taxa_Transferencia", TAXA_TRANSFERENCIA)
	Call GerenciadorGravaCampo("Incidente", "OK- Teste finalizou com sucesso")
End If


'******************************* VOLTA PARA O ORÇAMENTO***************************************************************************************
'********************************************************************************************************************************************
If vezes_seguradora = 1 Then
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("VOLTAR_ORCAMENTO")
	If Clicar(obj) = False Then ExitAction(False) End If
End If

ExitAction(True)
