VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAgregacaoD730"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DTO")
Option Explicit

' ==============================================================================
' DTO PARA AGREGAÇÃO DO REGISTRO D730
' Serve apenas para transportar dados em memória durante o processamento.
' ==============================================================================

Public ChaveAgrupamento As String
Public CST_ICMS As String
Public CFOP As String
Public ALIQ_ICMS As String
Public VL_OPR As Double
Public VL_BC_ICMS As Double
Public VL_ICMS As Double
Public VL_FCP As Double
