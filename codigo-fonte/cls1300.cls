VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls1300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ImportarLivroMovimentacaoCombustiveis()

Dim Arqs As Variant, Arq, Chave
Dim arrDados As New ArrayList
Dim versao As String
Dim i As Byte
    
    Arqs = Util.SelecionarArquivos("txt", "Selecione o SPED com o LMC")
    If VarType(Arqs) = vbBoolean Then Exit Function
    
    For Each Arq In Arqs
        
        If fnSPED.ValidarSPEDFiscal(Arq) Then Call fnSPED.ExtrairDadosLMC(Arq)
        
    Next Arq
    
    For Each Chave In dicRegistros.Keys()
        
        Select Case Chave
            
            Case "1300", "1310", "1320"
                Set arrDados = dicRegistros(Chave)
                Call Util.ExportarDadosArrayList(Worksheets(Chave), arrDados)
                
        End Select
        
    Next Chave
    
End Function

