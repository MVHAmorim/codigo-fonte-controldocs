VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAgruparRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AgruparRegistros(): End Sub

Private Sub CarregarDados(): End Sub

Private Sub ProcessarAgrupamento(): End Sub

Private Function ObterChaveOperacao(ByRef Campos As Variant) As String: End Function

Private Function AtualizarRegistroOperacao(ByRef Campos As Variant, ByVal Chave As String) As Variant: End Function
