VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncoesSPEDContribuicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ListarSPEDsContribuicoes(ByVal Arqs As Variant, ByRef arrContribuicoes As ArrayList)

Dim Arq As Variant
    
    For Each Arq In Arqs
        
        If fnSPED.ClassificarSPED(Arq) = "Fiscal" Then arrContribuicoes.Add Arq
        
    Next Arq
    
End Sub
