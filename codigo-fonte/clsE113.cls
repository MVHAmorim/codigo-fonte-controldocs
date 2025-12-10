VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsE113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CriarAtualizarE113(ByRef dicE113 As Dictionary, ByVal cAjuste, ByVal vAjuste As Double, _
                                   ByVal cPart As String, ByVal Modelo As String, ByVal SERIE As String, _
                                   ByVal nNF As String, ByVal Emissao As String, ByVal cProd As String, ByVal chNFe As String)
   
    If Not dicE113.Exists(cAjuste) Then Set dicE113(cAjuste) = CreateObject("System.Collections.ArrayList")
    dicE113(cAjuste).Add Join(Array("", "E113", cPart, Modelo, SERIE, "", nNF, Emissao, cProd, vAjuste, chNFe, ""), "|")
    
End Function
