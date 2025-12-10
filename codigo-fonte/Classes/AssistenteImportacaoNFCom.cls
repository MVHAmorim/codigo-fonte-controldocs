VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssistenteImportacaoNFCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================================
' CLASSE: AssistenteImportacaoNFCom
' ARQUITETURA: SLA (Single Level of Abstraction) com Orquestradores e Especialistas
' AUTOR: Equipe de Engenharia ControlDocs
' =========================================================================================

Private assImportacao As New AssistenteImportacao

' =========================================================================================
' 1. CAMADA PÚBLICA (ORQUESTRADOR NÍVEL 1)
' =========================================================================================

Public Sub ImportarNFCom(ByVal TipoImportacao As String)
    On Error GoTo TratarErro

    Call CarregarDocumentos(TipoImportacao)
    Call ExecutarProcessamentoEmLote
    Call ExportarResultadosExcel
    Call FinalizarProcesso

    Exit Sub
TratarErro:
    Call Util.MsgInformativa("Erro crítico na importação: " & VBA.Err.Description, "Erro", VBA.Now)
End Sub

Public Sub ProcessarNFCom(ByVal XML As DOMDocument60)
    On Error GoTo TratarErroProcessamento

    Dim NoInfNFCom As IXMLDOMNode
    Dim ChavePai As String
    Dim DicionarioAgregacao As Scripting.Dictionary

    Set NoInfNFCom = ValidarEstrutura(XML)
    If NoInfNFCom Is Nothing Then Exit Sub

    Call PrepararAmbienteGlobal(XML)
    Call GarantirParticipante(XML, NoInfNFCom)
    
    ChavePai = GerarRegistroD700(NoInfNFCom)
    
    Set DicionarioAgregacao = AgruparItensD730(NoInfNFCom)
    
    Call PersistirRegistrosFilhos(DicionarioAgregacao, ChavePai)

    Exit Sub
TratarErroProcessamento:
    Debug.Print "Erro ao processar arquivo (" & DadosXML.Arquivo & "): " & VBA.Err.Description
End Sub

' =========================================================================================
' 2. CAMADA DE ORQUESTRAÇÃO DE LOTE (PRIVADA NÍVEL 2)
' =========================================================================================

Private Sub CarregarDocumentos(ByVal TipoImportacao As String)
    Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    
    If TipoImportacao = "Arquivo" Then
        DocsFiscais.arrNFCom.addRange DocsFiscais.arrTodos
    End If
End Sub

Private Sub ExecutarProcessamentoEmLote()
    Dim Indice As Long
    Dim TempoInicio As Double
    Dim XML As Variant
    
    If DocsFiscais.arrNFCom.Count = 0 Then Exit Sub
    
    Call Util.AtualizarBarraStatus("Iniciando importação de NFCom...")
    TempoInicio = VBA.Timer
    Indice = 0
    
    For Each XML In DocsFiscais.arrNFCom
        Call Util.AntiTravamento(Indice, 10, "Importando NFCom " & Indice + 1, DocsFiscais.arrNFCom.Count, TempoInicio)
        Call ProcessarArquivoUnico(XML)
        Indice = Indice + 1
    Next XML
End Sub

Private Sub ExportarResultadosExcel()
    Dim Exportador As New ExportadorRegistros
    Call Exportador.ExportarRegistros("D700", "D730", "D731")
End Sub

Private Sub FinalizarProcesso()
    Call Util.MsgInformativa("Importação NFCom concluída com sucesso!", "Sucesso", VBA.Now)
    Call Util.AtualizarBarraStatus(False)
End Sub

' =========================================================================================
' 3. CAMADA DE EXECUÇÃO UNITÁRIA (PRIVADA NÍVEL 3)
' =========================================================================================

Private Sub ProcessarArquivoUnico(ByVal CaminhoArquivo As Variant)
    Dim DocumentoXML As DOMDocument60
    
    Set DocumentoXML = fnXML.RemoverNamespaces(CaminhoArquivo)
    
    If Not DocumentoXML Is Nothing Then
        Call ProcessarNFCom(DocumentoXML)
    End If
End Sub

' =========================================================================================
' 4. CAMADA DE LÓGICA DE NEGÓCIO (ESPECIALISTAS)
' =========================================================================================

Private Function ValidarEstrutura(ByVal XML As DOMDocument60) As IXMLDOMNode
    Set ValidarEstrutura = XML.SelectSingleNode("//*[local-name()='infNFCom']")
End Function

Private Sub PrepararAmbienteGlobal(ByVal XML As DOMDocument60)
    DadosXML.Arquivo = Util.NomeArquivo(DadosXML.Arquivo)
    DadosXML.COD_MOD = "62"
End Sub

Private Sub GarantirParticipante(ByVal XML As DOMDocument60, ByVal NoRaiz As IXMLDOMNode)
    Dim CNPJ As String, CPF As String
    Dim NoDest As IXMLDOMNode

    Set NoDest = NoRaiz.SelectSingleNode("dest")
    If NoDest Is Nothing Then Exit Sub

    CNPJ = fnXML.ValidarTag(NoDest, "CNPJ")
    CPF = fnXML.ValidarTag(NoDest, "CPF")

    If CNPJ <> "" Then
        Call assImportacao.CriarRegistro0150(XML, CNPJ)
    ElseIf CPF <> "" Then
        Call assImportacao.CriarRegistro0150(XML, CPF)
    End If
End Sub

Private Function GerarRegistroD700(ByVal NoRaiz As IXMLDOMNode) As String
    Dim Registro As CamposRegD700
    Dim Chave As String
    
    With Registro
        .REG = "D700"
        .Arquivo = Util.NomeArquivo(DadosXML.Arquivo)
        .COD_MOD = "62"
        .COD_SIT = "00"
        .IND_OPER = "0"
        .IND_EMIT = "0"

        Call PreencherIdentificacaoD700(Registro, NoRaiz)
        Call PreencherValoresD700(Registro, NoRaiz)
        Call PreencherParticipanteD700(Registro, NoRaiz)
        
        Chave = fnXML.ValidarTag(NoRaiz, "chNFCom")
        If Chave = "" Then Chave = fnXML.ValidarTag(NoRaiz, "Id")
        .CHV_REG = fnSPED.GerarChaveRegistro(.Arquivo, "D700", Chave)
    End With
    
    Call SalvarD700NoGlobal(Registro)
    GerarRegistroD700 = Registro.CHV_REG
End Function

Private Sub PreencherIdentificacaoD700(ByRef REG As CamposRegD700, ByVal NoRaiz As IXMLDOMNode)
    Dim NoIde As IXMLDOMNode
    Set NoIde = NoRaiz.SelectSingleNode("ide")
    If NoIde Is Nothing Then Exit Sub
    
    REG.SER = fnXML.ValidarTag(NoIde, "serie")
    REG.NUM_DOC = fnXML.ValidarTag(NoIde, "nNF")
    REG.DT_DOC = fnXML.ValidarTag(NoIde, "dhEmi")
    REG.DT_E_S = REG.DT_DOC
End Sub

Private Sub PreencherValoresD700(ByRef REG As CamposRegD700, ByVal NoRaiz As IXMLDOMNode)
    Dim NoTotal As IXMLDOMNode
    Set NoTotal = NoRaiz.SelectSingleNode("total")
    If NoTotal Is Nothing Then Exit Sub
    
    REG.VL_DOC = VBA.CDbl(fnXML.ValidarValores(NoTotal, "vNF"))
    REG.VL_PIS = VBA.CDbl(fnXML.ValidarValores(NoTotal, "vPIS"))
    REG.VL_COFINS = VBA.CDbl(fnXML.ValidarValores(NoTotal, "vCOFINS"))
    REG.VL_SERV = VBA.CDbl(fnXML.ValidarValores(NoTotal, "vProd"))
End Sub

Private Sub PreencherParticipanteD700(ByRef REG As CamposRegD700, ByVal NoRaiz As IXMLDOMNode)
    Dim NoDest As IXMLDOMNode
    Set NoDest = NoRaiz.SelectSingleNode("dest")
    If NoDest Is Nothing Then Exit Sub
    
    Dim CNPJ As String
    CNPJ = fnXML.ValidarTag(NoDest, "CNPJ")
    
    If CNPJ <> "" Then
        REG.COD_PART = CNPJ
    Else
        REG.COD_PART = fnXML.ValidarTag(NoDest, "CPF")
    End If
End Sub

Private Sub SalvarD700NoGlobal(ByRef REG As CamposRegD700)
    Dim Dados As Variant
    Dados = Array(REG.REG, REG.Arquivo, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_Fiscal, _
                  REG.CHV_PAI_CONTRIBUICOES, REG.IND_OPER, REG.IND_EMIT, REG.COD_PART, _
                  REG.COD_MOD, REG.COD_SIT, REG.SER, REG.NUM_DOC, REG.DT_DOC, _
                  REG.DT_E_S, REG.VL_DOC, REG.VL_DESC, REG.VL_SERV, REG.VL_SERV_NT, _
                  REG.VL_TERC, REG.VL_DA, REG.VL_BC_ICMS, REG.VL_ICMS, REG.COD_INF, _
                  REG.VL_PIS, REG.VL_COFINS, REG.CHV_DOCE, REG.FIN_DOCE, REG.TIP_FAT, _
                  REG.COD_MOD_DOC_REF, REG.CHV_DOCE_REF, REG.HASH_DOC_REF, REG.SER_DOC_REF, _
                  REG.NUM_DOC_REF, REG.MES_DOC_REF, REG.COD_MUN_DEST, REG.DED)

    If Not dtoRegSPED.rD700.Exists(REG.CHV_REG) Then
        dtoRegSPED.rD700.Add REG.CHV_REG, Dados
    End If
End Sub

Private Function AgruparItensD730(ByVal NoRaiz As IXMLDOMNode) As Scripting.Dictionary
    Dim dic As New Scripting.Dictionary
    Dim ListaDet As IXMLDOMNodeList
    Dim NoDet As IXMLDOMNode
    
    Set ListaDet = NoRaiz.SelectNodes("det")
    
    For Each NoDet In ListaDet
        Call ProcessarItemIndividual(NoDet, dic)
    Next NoDet
    
    Set AgruparItensD730 = dic
End Function

Private Sub ProcessarItemIndividual(ByVal NoDet As IXMLDOMNode, ByRef dic As Scripting.Dictionary)
    Dim DTO As clsAgregacaoD730
    Dim Chave As String
    
    Chave = GerarChaveAgrupamento(NoDet)
    
    If dic.Exists(Chave) Then
        Set DTO = dic(Chave)
    Else
        Set DTO = CriarNovoAgregado(Chave)
        dic.Add Chave, DTO
    End If
    
    Call AcumularValoresItem(DTO, NoDet)
End Sub

Private Function GerarChaveAgrupamento(ByVal NoDet As IXMLDOMNode) As String
    Dim NoImposto As IXMLDOMNode, NoProd As IXMLDOMNode
    Dim CST As String, Cfop As String, ALIQ As String
    
    Set NoImposto = NoDet.SelectSingleNode("imposto")
    Set NoProd = NoDet.SelectSingleNode("prod")
    
    CST = fnXML.ValidarTag(NoDet, "CST")
    If CST = "" And Not NoImposto Is Nothing Then CST = fnXML.ValidarTag(NoImposto, "CST")
    
    Cfop = fnXML.ValidarTag(NoDet, "CFOP")
    If Cfop = "" And Not NoProd Is Nothing Then Cfop = fnXML.ValidarTag(NoProd, "CFOP")
    
    ALIQ = "0"
    If Not NoImposto Is Nothing Then ALIQ = fnXML.ValidarValores(NoImposto, "pICMS")
    
    GerarChaveAgrupamento = CST & "|" & Cfop & "|" & ALIQ
End Function

Private Function CriarNovoAgregado(ByVal Chave As String) As clsAgregacaoD730
    Dim DTO As New clsAgregacaoD730
    Dim Partes() As String
    
    Partes = VBA.Split(Chave, "|")
    
    DTO.ChaveAgrupamento = Chave
    DTO.CST_ICMS = Partes(0)
    DTO.Cfop = Partes(1)
    DTO.ALIQ_ICMS = Partes(2)
    
    Set CriarNovoAgregado = DTO
End Function

Private Sub AcumularValoresItem(ByRef DTO As clsAgregacaoD730, ByVal NoDet As IXMLDOMNode)
    Dim NoImposto As IXMLDOMNode
    Set NoImposto = NoDet.SelectSingleNode("imposto")
    
    DTO.VL_OPR = DTO.VL_OPR + VBA.CDbl(fnXML.ValidarValores(NoDet, "vProd"))
    
    If Not NoImposto Is Nothing Then
        DTO.VL_BC_ICMS = DTO.VL_BC_ICMS + VBA.CDbl(fnXML.ValidarValores(NoImposto, "vBC"))
        DTO.VL_ICMS = DTO.VL_ICMS + VBA.CDbl(fnXML.ValidarValores(NoImposto, "vICMS"))
        DTO.VL_FCP = DTO.VL_FCP + VBA.CDbl(fnXML.ValidarValores(NoImposto, "vFCP"))
    End If
End Sub

Private Sub PersistirRegistrosFilhos(ByVal dic As Scripting.Dictionary, ByVal ChavePai As String)
    Dim Chave As Variant
    Dim DTO As clsAgregacaoD730
    
    For Each Chave In dic.Keys
        Set DTO = dic(Chave)
        Call CriarEGravarD730(DTO, ChavePai)
    Next Chave
End Sub

Private Sub CriarEGravarD730(ByVal DTO As clsAgregacaoD730, ByVal ChavePai As String)
    Dim REG As CamposRegD730
    
    With REG
        .REG = "D730"
        .Arquivo = Util.NomeArquivo(DadosXML.Arquivo)
        .CHV_PAI = ChavePai
        .CST_ICMS = DTO.CST_ICMS
        .Cfop = DTO.Cfop
        .ALIQ_ICMS = DTO.ALIQ_ICMS
        
        .VL_OPR = VBA.Format(DTO.VL_OPR, "0.00")
        .VL_BC_ICMS = VBA.Format(DTO.VL_BC_ICMS, "0.00")
        .VL_ICMS = VBA.Format(DTO.VL_ICMS, "0.00")
        .VL_RED_BC = "0,00"
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D730", .CST_ICMS, .Cfop, .ALIQ_ICMS)
    End With
    
    Call SalvarD730NoGlobal(REG)
    
    If DTO.VL_FCP > 0 Then
        Call CriarEGravarD731(REG, DTO.VL_FCP)
    End If
End Sub

Private Sub SalvarD730NoGlobal(ByRef REG As CamposRegD730)
    Dim Dados As Variant
    Dados = Array(REG.REG, REG.Arquivo, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_Fiscal, _
                  REG.CHV_PAI_CONTRIBUICOES, REG.CST_ICMS, REG.Cfop, REG.ALIQ_ICMS, _
                  REG.VL_OPR, REG.VL_BC_ICMS, REG.VL_ICMS, REG.VL_RED_BC, REG.COD_OBS)
                  
    If Not dtoRegSPED.rD730.Exists(REG.CHV_REG) Then
        dtoRegSPED.rD730.Add REG.CHV_REG, Dados
    End If
End Sub

Private Sub CriarEGravarD731(ByRef PaiD730 As CamposRegD730, ByVal ValorFCP As Double)
    Dim REG As CamposRegD731
    
    With REG
        .REG = "D731"
        .Arquivo = PaiD730.Arquivo
        .CHV_PAI = PaiD730.CHV_REG
        .VL_FCP_OP = VBA.Format(ValorFCP, "0.00")
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D731")
    End With
    
    Call SalvarD731NoGlobal(REG)
End Sub

Private Sub SalvarD731NoGlobal(ByRef REG As CamposRegD731)
    Dim Dados As Variant
    Dados = Array(REG.REG, REG.Arquivo, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_Fiscal, _
                  REG.CHV_PAI_CONTRIBUICOES, REG.VL_FCP_OP)

    If Not dtoRegSPED.rD731.Exists(REG.CHV_REG) Then
        dtoRegSPED.rD731.Add REG.CHV_REG, Dados
    End If
End Sub
