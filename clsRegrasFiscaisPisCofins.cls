VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegrasFiscaisPisCofins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ValidarCSTCreditosPresumidos(ByVal CST As Integer) As Boolean
    
    Select Case True
        
        Case CST Like "6*"
            ValidarCSTCreditosPresumidos = True
            
    End Select
    
End Function

Public Function ValidarCSTSemDireitoCredito(ByVal CST As Integer) As Boolean
    
    Select Case True
        
        Case CST Like "7*"
            ValidarCSTSemDireitoCredito = True
            
    End Select
    
End Function

Public Function ValidarCSTComDireitoCredito(ByVal CST As Integer) As Boolean
    
    Select Case True
        
        Case CST Like "5*"
            ValidarCSTComDireitoCredito = True
            
    End Select
    
End Function
