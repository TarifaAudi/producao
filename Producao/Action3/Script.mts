'************************************************CPF*********************************************************
'************************************************************************************************************
Set obj = Browser("Corretor Online - Porto").Page("Corretor Online - Porto").WebEdit("USUARIO_CPF")
If CampoSetSel(obj,"224.422.833-95") = False Then ExitAction(False) End If
'************************************************SENHA*******************************************************
'************************************************************************************************************
wait(2)
Set obj = Browser("Corretor Online - Porto").Page("Corretor Online - Porto").WebEdit("SENHA")
If CampoSetSel(obj,"@Auditeste8") = False Then ExitAction(False) End If


'**********************************************CHECKBOX CAPTCHA**********************************************
'************************************************************************************************************
Set obj = Browser("Corretor Online - Porto").Page("Corretor Online - Porto_2").Frame("Frame").WebCheckBox("CHKBOX_CAPTCHA")
If Clicar(obj) =  False Then ExitAction(False) End If
ExitAction(True)

'**********************************************BOTAO ENTRAR**************************************************
'************************************************************************************************************
Set obj = Browser("Corretor Online - Porto").Page("Corretor Online - Porto").WebButton("BTN_ENTRAR")
If Clicar(obj) = False Then ExitAction(False) End If
ExitAction(True)
