
'****************************************CLICA NO LINK AUTOMÓVEL NO COL**************************************
'************************************************************************************************************


Set obj = Browser("Browser").Page("Page").WebTable("Sessão duplicada Por motivos") 'FECHAR CASO ERRO 
If obj.Exist(0) Then
	Browser("Browser").Close
End If
Set obj = Browser("Porto Print - Porto Seguro").Page("Orçamento - Porto Print").WebElement("Ocorreu um erro na aplicação.")
If obj.Exist(0) Then
	If obj.GetROProperty("width") > 0 Then
		Browser("Porto Print - Porto Seguro").Page("Orçamento - Porto Print").WebButton("Sair").Click
		wait 2
		Browser("Porto Print - Porto Seguro").Close
	End If
End If

Set obj = Browser("name:=Porto Print.*").Page("title:=Porto Print.*").WebButton("html id:=bt_confirma_sair")
If obj.Exist(0) Then
	obj.Click
	wait 3
	If Browser("name:=Porto Print.*").Exist(0) Then
		Browser("name:=Porto Print.*").Close
	End If
	Set obj = Browser("title:=https://wwws\.portoseguro\.com.*")
	If obj.Exist(0) Then
		obj.Close
	End If
End If

Set obj = Browser("title:=https://wwws\.portoseguro\.com.*")
If obj.Exist(0) Then
	obj.Close
End If

If vezes_seguradora = 1 Then

'Set obj = Browser("title:=Porto Seguro.*").Page("title:=Porto Seguro.*").WebElement("innerhtml:=favoritos")
Set obj = Browser("Porto Seguro").Page("Porto Seguro").WebElement("favoritos")
If obj.Exist Then
	obj.Click
End If

Set objectlocation = Description.Create

objectlocation("html tag").value = "A"
objectlocation("micclass").value="Link"
objectlocation("innerhtml").value="Porto Print Web - Automóvel "

set obj = Browser("Micclass:=Browser").Page("Micclass:=Page").ChildObjects(objectlocation)
setlocation= obj.count

For i=3 to setlocation-1
	obj(i).click
	DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
	wait(2)
	Exit For
Next

'****************************************VERIFICA PÁGINA*****************************************************
'************************************************************************************************************
If VerificaPrimeiraPagina = False Then
	ExitAction(False)
End If

'Call VerificaPrimeiraPagina()



'************************************************VERIFICAÇÃO DE ARMAZENAMENTO DE TEMPO***********************
'************************************************************************************************************

i = 0
Do
	'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").WebElement("SELECAO_SUSEP_VERIFICACAO")
	Set obj = Browser("title:=Porto Print.*").Page("title:=Porto Print.*").WebList("name:=formIndex:susepLider")
	obj.RefreshObject
	If obj.Exist(0) Then
		DataAtual2 =  Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)
		Call ObtemHoraAtual(DataAtual, DataAtual2)
		Time_Portal_Apol_Renov = tempo
		DataAtual = Empty
		DataAtual2 = Empty
		Exit Do
	End If
	
	If i > 180 Then
		INCIDENTE = "Tela para Link Apolice/Renovação não abriu corretamente"
		GerenciadorGravaCampo "Incidente", INCIDENTE
		GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
		ErroAction = True
		ExitAction(False)
		Exit Do
	End If
	
	wait(1)
	i = i + 1
Loop


End If 



'****************************************LINK APOLICE RENOVACAO**********************************************
'************************************************************************************************************

Set obj = Browser("title:=Porto Print.*").Page("title:=Porto Print.*").Link("html id:=formHeaderEndosso:novoOrcamento")
'Set obj = Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("APOLICE_RENOVACAO") @@ hightlight id_;_Browser("Porto Print - Porto Seguro").Page("Porto Print - Porto Seguro").Link("Apólice/Renovação")_;_script infofile_;_ZIP::ssf1.xml_;_
If Clicar(obj) = False Then 
	ExitAction(False) 
End If

DataAtual = Right("0" & Datepart("h",Now()),2) & ":" & Right("0" & Datepart("n",Now()),2) & ":" & Right("0" & Datepart("s",Now()),2)

ExitAction(True)



Function VerificaPrimeiraPagina()
	i = 0
	Do
		Set obj = Browser("title:=Porto Print.*").Page("title:=Porto Print.*").WebList("name:=formIndex:susepLider")
		obj.RefreshObject
		If obj.Exist(0) Then
			If obj.GetROProperty("width") > 1 Then
				VerificaPrimeiraPagina = True
				Exit Function
			End If
		End If
		
		'Set obj = Browser("Browser").Page("Page").WebTable("Sessão duplicada Por motivos")
		Set obj = Browser("title:=https://wwws\.portoseguro\.com.*").Page("name:=gtm_autoEvent.*").WebTable("html id:=tabela")
		obj.RefreshObject
		If obj.Exist(0) Then
			If obj.GetROProperty("width") > 1 Then
				Set objKey = CreateObject("WScript.Shell")
				objkey.SendKeys"{F6}"
				objkey.SendKeys"{F6}"
				objkey.SendKeys"{F6}"
				wait(1)
				objkey.SendKeys"https://wwws.portoseguro.com.br/portoprintweb/servlet/PPWFinalServletWithoutReturn?forceFinal=true"
				objkey.SendKeys"{ENTER}"
				wait(1)
				Set objectlocation = Description.Create
				
				objectlocation("html tag").value = "A"
				objectlocation("micclass").value="Link"
				objectlocation("innerhtml").value="Porto Print Web - Automóvel "
				
				set obj = Browser("Micclass:=Browser").Page("Micclass:=Page").ChildObjects(objectlocation)
				setlocation= obj.count
				
				For i = 3 to setlocation-1
					obj(i).click
					Exit For
				Next
			End If
		End If
		
		If i > 300 Then
			INCIDENTE = "Verificar Print da Tela, Erro Impedindo Execução do Robô!"
			GerenciadorGravaCampo "Incidente", INCIDENTE
			GerenciadorSalvaEvidenciaCapturaTela("Evidencia")
			ErroAction = True
			VerificaPrimeiraPagina = False
			Exit Function
		End If
	wait(0.2)
	i = i + 1
	Loop
	
End Function
