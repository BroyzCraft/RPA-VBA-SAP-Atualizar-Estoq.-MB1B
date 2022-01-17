Attribute VB_Name = "script"
Sub atualizarDadosEstoque()

login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    On Error GoTo sap_login
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If

Dim totalLinhas, primeiraLinha, ultimaLinha
Sheets("dados").Select
totalLinhas = Range("A10000").End(xlUp).Row
primeiraLinha = 2
ultimaLinha = 8
aux = 2

Sheets("menu").Select
depSaida = Range("B1").Value
depEntrada = Range("B2").Value

'REGISTRAR TRANSAÇÃO
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB1B"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "411"
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "1000"
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = depSaida
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").SetFocus
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").caretPosition = 4
    
Do While aux < totalLinhas + 1
    
    On Error GoTo tratamentoErro
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").Text = depEntrada
    session.findById("wnd[0]").sendVKey 0
    Range("A1:B12").Select
    
    Sheets("dados").Select
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = Range("A" & primeiraLinha).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[1,7]").Text = Range("A" & primeiraLinha + 1).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[2,7]").Text = Range("A" & primeiraLinha + 2).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[3,7]").Text = Range("A" & primeiraLinha + 3).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[4,7]").Text = Range("A" & primeiraLinha + 4).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[5,7]").Text = Range("A" & primeiraLinha + 5).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[6,7]").Text = Range("A" & primeiraLinha + 6).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[7,7]").Text = Range("A" & primeiraLinha + 7).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[8,7]").Text = Range("A" & primeiraLinha + 8).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[9,7]").Text = Range("A" & primeiraLinha + 9).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[10,7]").Text = Range("A" & primeiraLinha + 10).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[11,7]").Text = Range("A" & primeiraLinha + 11).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[12,7]").Text = Range("A" & primeiraLinha + 12).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[13,7]").Text = Range("A" & primeiraLinha + 13).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[14,7]").Text = Range("A" & primeiraLinha + 14).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[15,7]").Text = Range("A" & primeiraLinha + 15).Value
    
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = Range("B" & primeiraLinha).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[1,26]").Text = Range("B" & primeiraLinha + 1).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[2,26]").Text = Range("B" & primeiraLinha + 2).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[3,26]").Text = Range("B" & primeiraLinha + 3).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[4,26]").Text = Range("B" & primeiraLinha + 4).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[5,26]").Text = Range("B" & primeiraLinha + 5).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[6,26]").Text = Range("B" & primeiraLinha + 6).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[7,26]").Text = Range("B" & primeiraLinha + 7).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[8,26]").Text = Range("B" & primeiraLinha + 8).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[9,26]").Text = Range("B" & primeiraLinha + 9).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[10,26]").Text = Range("B" & primeiraLinha + 10).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[11,26]").Text = Range("B" & primeiraLinha + 11).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[12,26]").Text = Range("B" & primeiraLinha + 12).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[13,26]").Text = Range("B" & primeiraLinha + 13).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[14,26]").Text = Range("B" & primeiraLinha + 14).Value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[15,26]").Text = Range("B" & primeiraLinha + 15).Value
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    aux = aux + 16
    
    'cria backup
    Range("A" & primeiraLinha & ":B" & primeiraLinha + 10).Copy
    tot = Range("G1000000").End(xlUp).Row
    Range("G" & tot + 1).PasteSpecial xlPasteValues
    
    'apaga
    Range("A" & primeiraLinha & ":B" & primeiraLinha + 10).Select
    Selection.Delete Shift:=xlUp

Loop

MsgBox ("Atualizações Realizadas!")
End

tratamentoErro:
    MsgBox ("Erro dos registros!")
    End
    
End Sub

Sub atualizarDadosEstoqueUnitario()

login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    On Error GoTo sap_login
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If

Dim totalLinhas, primeiraLinha, ultimaLinha
Sheets("dados").Select
totalLinhas = Range("A10000").End(xlUp).Row
primeiraLinha = 2
ultimaLinha = 8
aux = 2

Sheets("menu").Select
depSaida = Range("B1").Value
depEntrada = Range("B2").Value

'REGISTRAR TRANSAÇÃO
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB1B"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "411"
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "1000"
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = depSaida
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").SetFocus
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").caretPosition = 4
    
Do While aux < totalLinhas + 1
    
    On Error GoTo tratamentoErro
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").Text = depEntrada
    session.findById("wnd[0]").sendVKey 0
    
    Sheets("dados").Select
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = Range("A" & primeiraLinha).Value
    
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = Range("B" & primeiraLinha).Value
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    aux = aux + 2
    
    'cria backup
    Range("A" & primeiraLinha & ":B" & primeiraLinha).Copy
    tot = Range("G1000000").End(xlUp).Row
    Range("G" & tot + 1).PasteSpecial xlPasteValues
    
    'apaga
    Range("A" & primeiraLinha & ":B" & primeiraLinha).Select
    Selection.Delete Shift:=xlUp

Loop

MsgBox ("Atualizações Realizadas!")
End

tratamentoErro:
    tot = Range("G1000000").End(xlUp).Row
    Range("I" & tot).Value = "X"
    MsgBox ("Erro dos registros!")
    End
    
End Sub

