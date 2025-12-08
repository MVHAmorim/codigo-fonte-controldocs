VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTratamentoErros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TratarErros(ByRef Erro As ErrObject, ByVal nRotina As String)

Dim Msg As String, Tabela As String
    
    Select Case Erro.Number
            
        Case Else
            Msg = "Erro sem tratamento na rotina: " & nRotina & vbCrLf & vbCrLf
            Msg = Msg & "Código do erro: " & Err.Number & vbCrLf
            Msg = Msg & "Descrição do erro: " & Err.Description
            
            MsgBox Msg, vbCritical, "Erro não identificado"
            
    End Select
    
End Sub
