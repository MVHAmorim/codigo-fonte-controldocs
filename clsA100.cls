VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsA100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ImpPlanilha As New clsA100_ImportadorPlanilha

Public Sub ImportarPlanilhaA100_A170()
    
    Call ImpPlanilha.ImportarPlanilha
    
End Sub

Private Sub InicializarObjetos()
    
    Set ImpPlanilha = New clsA100_ImportadorPlanilha
    
End Sub

Private Sub LimparObjetos()
    
    Set ImpPlanilha = Nothing
    
End Sub
