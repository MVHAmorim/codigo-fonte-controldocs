VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssistentesInteligentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Fiscal As New clsAssistentesFiscais
Public Estoque As New clsAssistenteEstoque
Public Tributacao As New clsAssistenteTributacao
Public Tributario As New AssistenteTributario

Public Function ResetarInconsistencias()

Dim Resposta As VbMsgBoxResult
    
    Resposta = MsgBox("Tem certeza que deseja resetar TODAS as inconsistências?" & vbCrLf & _
        "Essa operação NÃO pode ser desfeita.", vbExclamation + vbYesNo, "Resetar Inconsistências")
        
    If Resposta = vbNo Then Exit Function
    
    Call dicInconsistenciasIgnoradas.RemoveAll
    Call Util.MsgAviso("Todas as Inconsistências foram resetadas com sucesso!", "Reset de Inconsistências")
    
    With ActiveSheet
        
        Select Case .CodeName
            
            Case "assApuracaoICMS"
                Call Assistente.Fiscal.Apuracao.ICMS.ReprocessarSugestoes
                
            Case "assApuracaoIPI"
                Call Assistente.Fiscal.Apuracao.IPI.ReprocessarSugestoes
                
            Case "assApuracaoPISCOFINS"
                Call Assistente.Fiscal.Apuracao.PISCOFINS.ReprocessarSugestoes
                
            Case "relInteligenteDivergencias"
                Call Assistente.Fiscal.Divergencias.ReprocessarSugestoes
                
            Case "relInteligenteEstoque"
                Call Assistente.Estoque.ReprocessarSugestoes
        
        End Select
        
        .Activate
        
    End With
    
End Function

