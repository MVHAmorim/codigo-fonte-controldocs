VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncoesSPEDFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ImportarSPEDFiscal(Optional ByVal SelReg As String, Optional Periodo As String, Optional Unificar As Boolean)



End Function

Public Sub ListarSPEDsFiscais(ByVal Arqs As Variant, ByRef arrFiscal As ArrayList)

Dim Arq As Variant
    
    For Each Arq In Arqs
        
        If fnSPED.ClassificarSPED(Arq) = "Fiscal" Then arrFiscal.Add Arq
        
    Next Arq
    
End Sub
