Attribute VB_Name = "Variaveis"
Option Explicit
' SETAR SAP
Public SapGuiAuto As Variant
Public SAPApp As Variant
Public SAPCon As Variant
Public session As Variant
Public Connection As Variant
Public WScript As Variant

' SETAR PLANILHA
Public SheetNumber As Integer
Public ThisWorkbook As Workbook
Public ThisWorksheet As Worksheet
Public ListaFornecedores As Worksheet

' VARIAVEIS DINAMICAS NO LOOP DA PLANILHA
Public Ordem As String
Public DataReal As String

' VARIAVEIS DINAMICAS NO LOOP DOS COMPONENTES SAP
Public MaterialFaltante As String
Public DataPlanejada As String
Public SecaoCausadora As String
Public Projeto As String
Public DescricaoMaterial As String
Public Fornecedor As String
Public StatusComponente As String
Public numero_ate_onde_operacao_vai As Integer

Sub DeclararVariaveis(sheetNum)
    SheetNumber = sheetNum
    
    Set ThisWorkbook = ActiveWorkbook
    Set ThisWorksheet = ThisWorkbook.Worksheets(sheetNum)
    Set ListaFornecedores = ThisWorkbook.Worksheets(4)
    
    Select Case SheetNumber
        Case Is = 1
            numero_ate_onde_operacao_vai = 9999
        Case Is = 2
            numero_ate_onde_operacao_vai = 850
        Case Is = 3
            numero_ate_onde_operacao_vai = 1250
    End Select
End Sub

Sub ConectarSAP()
    On Error GoTo AlertaAbrirSAP
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPApp = SapGuiAuto.GetScriptingEngine
        Set SAPCon = SAPApp.Children(0)
        Set session = SAPCon.Children(0)
        If Not IsObject(Application) Then
            Set SapGuiAuto = GetObject("SAPGUI")
        End If
        If Not IsObject(session) Then
            Set session = Connection.Children(0)
        End If
        If IsObject(WScript) Then
            WScript.ConnectObject session, "on"
            WScript.ConnectObject Application, "on"
        End If
        Exit Sub
AlertaAbrirSAP:
    MsgBox "Abra o SAP para rodar a macro"
End Sub
