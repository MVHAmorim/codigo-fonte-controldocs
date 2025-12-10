VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGerenciadorErros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub NotificarErroInesperado(ByVal NomeClasse As String, ByVal NomeMetodo As String, Optional MensagemAdicional As String = "")
    
    With infNotificacao
        
        .Classe = NomeClasse
        .Funcao = NomeMetodo
        .MensagemErro = Err.Number & " - " & Err.Description
        .OBSERVACOES = MensagemAdicional
        
    End With
    
    Call Notificacoes.NotificarErroInesperado
    
End Sub
