VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AtualizadorTabelasSPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub BaixarTabela(ByVal URL As String, ByVal NomeTabela As String)

Dim http As New XMLHTTP60
Dim Dados As String
    
    http.Open "GET", URL, False
    http.Send
    
    If http.Status = 200 Then
        
        Dados = http.ResponseText
        Call SalvarTabela(Dados, NomeTabela)
        
    End If
    
End Sub

Private Function SalvarTabela(ByVal Dados As String, ByVal NomeTabela As String)

Dim CustomPart As New clsCustomPartXML
    
    Call CustomPart.SalvarTabelaTXT(NomeTabela, Dados)
    
End Function

Public Sub AtualizarTabelas()

    

End Sub
