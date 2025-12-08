VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fnSPED_AgruparRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Iprocessador

Public TipoSPED As String
Public Registro As String

Public Sub IProcessador_Executar(ByVal tpSPED As String, ByVal nReg As String)
    
    Me.TipoSPED = tpSPED
    Me.Registro = nReg
    
    Call AgruparRegistros
    
End Sub

Private Sub AgruparRegistros()

Dim Agrupador As IAgruparRegistros
    
    Select Case True
        
        Case TipoSPED = "Contribuições" And Registro = "C180"
            Set Agrupador = New fnSPED_AgruparRegistrosC180Cont
            Call Agrupador.AgruparRegistros
            
        Case TipoSPED = "Contribuições" And Registro = "C190"
            Set Agrupador = New fnSPED_AgruparRegistrosC190Cont
            Call Agrupador.AgruparRegistros
            
        Case Else
            Call Util.MsgAlerta("Registro não mapeado para função de Agrupamento de registros", "Registro Não Mapeado")
            
    End Select
    
End Sub
