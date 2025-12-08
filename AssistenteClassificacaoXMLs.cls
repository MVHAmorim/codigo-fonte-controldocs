VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssistenteClassificacaoXMLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ClassificadorArquivos As New AssistenteClassificacaoArquivos

Public Function ListarArquivosXML(ByVal Caminho As String)

Dim ArquivosListados As Variant
    
    On Error GoTo Notificar:
        
    If Not ClassificadorArquivos.EncontrouPowerShell() Then Exit Function
    
    Call Util.AtualizarBarraStatus("Listando arquivos selecionados...")
    ArquivosListados = ClassificadorArquivos.ListarArquivos(Caminho, "xml")
    
    Call ClassificarXMLs(ArquivosListados)
    
Exit Function
Notificar:

    Call TratarExcecoes
    
End Function

Public Function ClassificarXMLs(ByVal ArquivosListados As Variant)

Dim Doce As New DOMDocument60
Dim TotalArquivos As Long
Dim XML As Variant
    
    On Error GoTo Notificar:
    
    Doce.async = False
    Doce.validateOnParse = False
    
    ArquivosListados = Util.ConverterArrayListEmArray(ArquivosListados)
    
    a = 0
    Comeco = Timer()
    TotalArquivos = UBound(ArquivosListados) + 1
    For Each XML In ArquivosListados
        
        Call Util.AntiTravamento(a, 50, "Classificando XMLs, por favor aguarde...", TotalArquivos, Comeco)
        
        Set Doce = fnXML.RemoverNamespaces(XML)
        Call ClassificarXML(Doce, XML)
        
    Next XML
    
Exit Function
Notificar:

    Call TratarExcecoes
    
End Function

Private Function ClassificarXML(ByRef Doce As DOMDocument60, ByVal Arquivo As String)
    
    With DocsFiscais
        
        Select Case True
            
            Case Not fnXML.ValidarXML(Doce)
                If Not .arrDocsInvalidos.contains(Arquivo) Then .arrDocsInvalidos.Add Arquivo
                
            Case fnXML.ValidarNFe(Doce)
                If Not .arrNFeNFCe.contains(Arquivo) Then .arrNFeNFCe.Add Arquivo
                
            Case fnXML.ValidarCTe(Doce)
                If Not .arrCTe.contains(Arquivo) Then .arrCTe.Add Arquivo
                
            Case fnXML.ValidarCFe(Doce)
                If Not .arrCFe.contains(Arquivo) Then .arrCFe.Add Arquivo
                
            Case fnXML.ValidarNFSe(Doce)
                If Not .arrNFSe.contains(Arquivo) Then .arrNFSe.Add Arquivo
                
            Case fnXML.ValidarProtocoloCancelamento(Doce)
                If Not .arrProtocolos.contains(Arquivo) Then .arrProtocolos.Add Arquivo
                
            Case Else
                If Not .arrDocsInvalidos.contains(Arquivo) Then .arrDocsInvalidos.Add Arquivo
                
        End Select
        
    End With
    
    Set Doce = Nothing
    
End Function

Private Function TratarExcecoes()

    Select Case True

        Case Else
            With infNotificacao
        
                .Funcao = "ListarArquivosSPED"
                .Classe = "AssistenteClassificacaoSPEDs"
                .MensagemErro = Err.Number & " - " & Err.Description
                .OBSERVACOES = "Erro Inesperado"
                
            End With
            
            Call Notificacoes.NotificarErroInesperado
            
    End Select

End Function
