'************************************************VERIFICAÇÃO DE ABA ATIVA************************************
'************************************************************************************************************

i = 0 
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("ABA_QUESTIONARIO")
obj.RefreshObject
If obj.Exist(0) Then
	If obj.GetROProperty("class") = "aba-ativa" Then
		DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
		Call ObtemHoraAtual(DataAtual, DataAtual2)
		Time_Veiculo_Questionario = tempo
		DataAtual = Empty
		DataAtual2 = Empty
		Exit Do
	End If
End If

If i > 180 Then
	INCIDENTE = "Aba Questionário Não Carregou no tempo limite" & i & "!"
	GerenciadorGravaCampo "Incidente", INCIDENTE
	GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
	ErroAction = True
	ExitAction(False)
	Exit Do
End If

wait(1)
i = i + 1
Loop


'*******************************************NOME DO PROPONENTE***********************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("NOME_PROPONENTE")
If CampoSetSel(obj,"teste teste") = False Then ExitAction(False) End If

'*******************************RELAÇÃO DO PRINCIPAL CONDUTOR COM SEGURADO***********************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("RELACAO_CONDUTOR_SEGURADO")
obj.Click
Set objKey = CreateObject("WScript.Shell")
objkey.SendKeys"{DOWN}"
objkey.SendKeys"{ENTER}"
wait(2)
'*******************************CPF**************************************************************************
'************************************************************************************************************
i = 0
Do
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CPF")
obj.RefreshObject
If obj.GetROProperty("disabled") = 1 Then
	Exit Do
Else
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("RELACAO_CONDUTOR_SEGURADO")
	If CampoSetSel(obj,"Filho(a)") = False Then ExitAction(False) End If
	wait(2)
	Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("RELACAO_CONDUTOR_SEGURADO")
	If CampoSetSel(obj,"O próprio") = False Then ExitAction(False) End If
End If

If i > 15 Then
	INCIDENTE = "Relação Principal condutor Segurado não foi preenchido corretamente, campo O próprio não desabilitando cpf, nome, data de nascimento"
	Exit Do
End If
wait(1)
i = i = 1
Loop


'************************************************SEXO********************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("SEXO")
If CampoSetSel(obj,"Masculino") = False Then ExitAction(False) End If
'************************************ESTADO CIVIL DO PRINCIPAL CONDUTOR**************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("ESTADO_CIVIL_PRINCIPAL_CONDUTOR")
If CampoSetSel(obj,"Solteiro(a)") = False Then ExitAction(False) End If
'***************RESIDEM COM O PRINCIPAL CONDUTOR, PESSOAS NA FAIXA ETÁRIA ENTRE 18 A 24 ANOS?****************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("NAO_E_ESTOU_PELANAMENTE_CIENTE")
If Clicar(obj) = False Then ExitAction(False) End If
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SIM_E_NAO_UTILIZAM_O_VEICULO")
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("SIM_E_UTILIZAM_O_VEICULO")
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("NAO_INFORMADO_PRINCIPAL_CONDUTOR")

'*************************************************RESIDE EM**************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("RESIDE_EM")
If CampoSetSel(obj,"Casa/Sobrado") = False Then ExitAction(False) End If
'*********************QUAL A DISTANCIA DA RESIDENCIA ATÉ O SEU LOCAL DE TRABALHO*****************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("QUAL_A_DISTANCIA_RESIDENCIA_LOCAL_TRABALHO")
If CampoSetSel(obj,"Até 20 km") = False Then ExitAction(False) End If
'*****************************QUAL A ATIVIDADE PROFISSIONAL QUE EXERCE?**************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("QUAL_ATIVIDADE_PROFISSIONAL_EXERCE")
If CampoSetSel(obj,"Dentista") = False Then ExitAction(False) End If
''*****************************************CEP NÚMERO****************************************************
''************************************************************************************************************
'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CEP_NUMERO")
'If CampoSetSel(obj,"") = False Then ExitAction(False) End If
''******************************************CEP COMPLEMENTO***************************************************
''************************************************************************************************************
'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebEdit("CEP_COMPLEMENTO")
'If CampoSetSel(obj,"Solteiro(a)") = False Then ExitAction(False) End If

'******************************************BOTAO COPIAR CEP PROPONENTE***************************************
'************************************************************************************************************
wait(2)
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("BTN_COPIAR_CEP")
If Clicar(obj) = False Then ExitAction(False) End If

'****************************************CHECKBOX CEP NÃO INFORMADO******************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("NAO_INFORMADO_CEP")


'********************************************NA RESIDÊNCIA***************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("NA_RESIDENCIA")
If CampoSetSel(obj,"Não, na residência") = False Then ExitAction(False) End If
'**********************************************NO TRABALHO***************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("NO_TRABALHO")
If CampoSetSel(obj,"Não, no trabalho") = False Then ExitAction(False) End If
'************************************NO COLÉGIO/FACULDADE/PÓS GRADUAÇÃO**************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("NO_COLEGIO_FACULDADE")
If CampoSetSel(obj,"Não estuda ou o veículo não é utilizado como meio de transporte ao colégio/faculdade/pós-graduação") = False Then ExitAction(False) End If
'***********************************************UTILIZA VEÍCULO**********************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("ULTILIZA_VEICULO_DOIS_OU_MAIS_DIAS_SIM")
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("UTILIZA_VEICULO_DOIS_OU_MAIS_DIAS_NAO")
If Clicar(obj) = False Then ExitAction(False) End If
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebCheckBox("UTILIZA_VEICULO_DOIS_OU_MAIS_DIAS_NAO_INFORMADO")

'********************************************POSSUI DISPOSITIVO ANTIFURTO************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebList("POSSUI_DISPOSITIVO_ANTIFURTO")
If CampoSetSel(obj,"DAF-V (Rastreador da Porto Seguro)") = False Then ExitAction(False) End If
wait(4)
'********************************************BOTÃO AVANÇAR***************************************************
'************************************************************************************************************
Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Continuar")
If Clicar(obj) = False Then ExitAction(False) End If

DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)

ExitAction(True)
