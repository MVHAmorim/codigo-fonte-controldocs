VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegrasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
        Option Explicit

Public Geral As New clsRegrasFiscaisGerais
Public PISCOFINS As New clsRegrasFiscaisPisCofins
Public ApuracaoICMS As New clsRegrasApuracaoICMS
Public ApuracaoIPI As New clsRegrasApuracaoIPI
Public ApuracaoPISCOFINS As New clsRegrasApuracaoPISCOFINS
Public SPEDFiscal As New clsRegrasSPEDFiscal
Public SPEDContribuicoes As New clsRegrasSPEDContribuicoes
Public DivergenciasFiscais As New clsRegrasDivergencias

