VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssistenteImportacaoNFCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Servicos\Importacao")
Option Explicit

' =========================================================================================
' CLASSE: AssistenteImportacaoNFCom
' ARQUITETURA: SLA (Single Level of Abstraction) - Refatorada
' AUTOR: Equipe de Engenharia ControlDocs
' =========================================================================================

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private assImportacao As New AssistenteImportacao
Private GerenciadorSPED As New clsRegistrosSPED

' Variáveis de Contexto da Classe
Private pDicionarioAgregacao As Scripting.Dictionary

Private Sub Class_Initialize()
    Set pDicionarioAgregacao = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set pDicionarioAgregacao = Nothing
End Sub

' =========================================================================================
' 1. ORQUESTRADOR PRINCIPAL (PÚBLICO)
' =========================================================================================

Public Sub ImportarNFCom(ByVal TipoImportacao As String)
    On Error GoTo TratarErro

    Call CarregarDocumentos(TipoImportacao)
    Call CarregarDadosExistentes
    Call ExecutarProcessamentoEmLote
    Call ExportarResultadosExcel
    Call FinalizarProcesso

    Exit Sub
TratarErro:
    Call Util.MsgInformativa("Erro crítico na importação: " & VBA.Err.Description, "Erro", VBA.Now)
    Call Util.AtualizarBarraStatus(False)
End Sub

' =========================================================================================
' 2. PROCESSAMENTO DE LOTE (PRIVADO)
' =========================================================================================

Private Sub CarregarDocumentos(ByVal TipoImportacao As String)
    Call ProcessarDocumentos.CarregarXMLS(TipoImportacao)
    If TipoImportacao = "Arquivo" Then
        DocsFiscais.arrNFCom.addRange DocsFiscais.arrTodos
    End If
End Sub

Private Sub CarregarDadosExistentes()
    Call Util.AtualizarBarraStatus("Carregando estrutura SPED existente...")
    
    With GerenciadorSPED
        ' Carrega hierarquia básica
        .CarregarDadosRegistro0000 "Arquivo"
        .CarregarDadosRegistro0000_Contr "Arquivo"
        .CarregarDadosRegistro0001 "Arquivo"
        
        ' Tenta carregar blocos que podem existir
        On Error Resume Next
        .CarregarDadosRegistroD001 "Arquivo"
        .CarregarDadosRegistroD700 "CHV_DOCE" ' Chave específica para facilitar busca
        On Error GoTo 0
    End With
End Sub

Private Sub ExecutarProcessamentoEmLote()
    Dim Indice As Long
    Dim TempoInicio As Double
    Dim XML As Variant
    
    If DocsFiscais.arrNFCom.Count = 0 Then Exit Sub
    
    Call Util.AtualizarBarraStatus("Iniciando processamento de NFCom...")
    TempoInicio = VBA.Timer
    Indice = 0
    
    For Each XML In DocsFiscais.arrNFCom
        Call Util.AntiTravamento(Indice, 10, "Processando NFCom " & Indice + 1, DocsFiscais.arrNFCom.Count, TempoInicio)
        Call ProcessarArquivoUnico(XML)
        Indice = Indice + 1
    Next XML
End Sub

Private Sub ProcessarArquivoUnico(ByVal CaminhoArquivo As Variant)
    Dim DocumentoXML As DOMDocument60
    Set DocumentoXML = fnXML.RemoverNamespaces(CaminhoArquivo)
    
    If Not DocumentoXML Is Nothing Then
        Call ProcessarNFCom(DocumentoXML)
    End If
End Sub

' =========================================================================================
' 3. PROCESSAMENTO DE DOCUMENTO (CORE LÓGICO)
' =========================================================================================

Private Sub ProcessarNFCom(ByRef XML As DOMDocument60)
    On Error GoTo ErroProc
    
    ' 1. Validação Básica
    If XML.SelectSingleNode("//*[local-name()='infNFCom']") Is Nothing Then Exit Sub
    
    ' 2. Preparação do Contexto
    Call PrepararAmbienteGlobal(XML)
    Call GarantirHierarquiaSPED(XML)
    Call GarantirParticipante(XML)
    
    ' 3. Geração dos Registros
    Call CriarRegistroD700(XML)
    
    Exit Sub
ErroProc:
    Debug.Print "Erro ao processar XML (" & DadosXML.ARQUIVO & "): " & VBA.Err.Description
End Sub

' =========================================================================================
' 4. PREPARAÇÃO DO AMBIENTE (CONTEXTO)
' =========================================================================================

Private Sub PrepararAmbienteGlobal(ByRef XML As DOMDocument60)
    Dim DataEmissao As String
    Dim Ano As String, Mes As String
    
    Call DTO_DadosXML.ResetarDadosXML
    
    With DadosXML
        .COD_MOD = "62"
        .CNPJ_EMITENTE = fnXML.ValidarTag(XML, "//emit/CNPJ")
        .CNPJ_DESTINATARIO = fnXML.ValidarTag(XML, "//dest/CNPJ")
        If .CNPJ_DESTINATARIO = "" Then .CNPJ_DESTINATARIO = fnXML.ValidarTag(XML, "//dest/CPF")
        
        .CNPJ_ESTABELECIMENTO = CNPJContribuinte
        
        ' Extração de Data para Período (Fundamental para registro 0000)
        DataEmissao = VBA.Left(fnXML.ValidarTag(XML, "//ide/dhEmi"), 10)
        
        If VBA.Len(DataEmissao) >= 10 Then
            Ano = VBA.Left(DataEmissao, 4)
            Mes = VBA.Mid(DataEmissao, 6, 2)
            .Periodo = Mes & "/" & Ano
        Else
            .Periodo = VBA.Format(Date, "mm/yyyy")
        End If
        
        .ARQUIVO = .Periodo & "-" & .CNPJ_ESTABELECIMENTO
    End With
End Sub

Private Sub GarantirHierarquiaSPED(ByRef XML As DOMDocument60)
    ' Verifica e cria Registro 0000 (Fiscal)
    If Not dtoRegSPED.r0000.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0000(XML)
    End If
    
    ' Verifica e cria Registro 0000 (Contribuições)
    If Not dtoRegSPED.r0000_Contr.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0000_Contr(XML)
    End If
    
    ' Verifica e cria Registro 0001 (Abertura Bloco 0)
    If Not dtoRegSPED.r0001.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0001
    End If
    
    ' Verifica e cria Registro D001 (Abertura Bloco D)
    If Not dtoRegSPED.rD001.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistroD001
    End If
End Sub

Private Sub GarantirParticipante(ByRef XML As DOMDocument60)
    Dim NoDest As IXMLDOMNode
    Dim CNPJ As String, CPF As String
    
    Set NoDest = XML.SelectSingleNode("//dest")
    If NoDest Is Nothing Then Exit Sub
    
    CNPJ = fnXML.ValidarTag(NoDest, "CNPJ")
    CPF = fnXML.ValidarTag(NoDest, "CPF")
    
    If CNPJ <> "" Then
        Call assImportacao.CriarRegistro0150(XML, CNPJ)
    ElseIf CPF <> "" Then
        Call assImportacao.CriarRegistro0150(XML, CPF)
    End If
End Sub

' =========================================================================================
' 5. GERAÇÃO DE REGISTROS (ESPECIALISTAS)
' =========================================================================================

Public Sub CriarRegistroD700(ByRef NFCom As DOMDocument60)
    Dim Campos As Variant
    Dim Chave As String
    Dim ChavePai As String
    
    ' Carrega Dicionário se necessário
    If dtoRegSPED.rD700 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD700("CHV_DOCE")
    
    With CamposD700
        ' Identificação e Controle
        .REG = "D700"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_MOD = "62"
        .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(fnXML.ValidarTag(NFCom, "//cStat"))))
        
        ' Chaves e Datas
        .SER = VBA.Format(fnXML.ValidarTag(NFCom, "//ide/serie"), "000")
        .NUM_DOC = fnXML.ValidarTag(NFCom, "//ide/nNF")
        .DT_DOC = fnXML.ValidarTag(NFCom, "//ide/dhEmi")
        .DT_E_S = .DT_DOC
        
        .CHV_DOCE = fnXML.ValidarTag(NFCom, "//chNFCom")
        If .CHV_DOCE = "" Then .CHV_DOCE = VBA.Replace(fnXML.ValidarTag(NFCom, "//infNFCom/@Id"), "NFCom", "")
        
        ' Participante
        .COD_PART = assImportacao.ExtrairCodigoParticipante()
        If .COD_PART <> "" Then .COD_PART = Util.FormatarTexto(.COD_PART)
        
        ' Valores Monetários
        .VL_DOC = fnXML.ValidarValores(NFCom, "//total/vNF")
        .VL_DESC = fnXML.ValidarValores(NFCom, "//total/vDesc")
        .VL_SERV = fnXML.ValidarValores(NFCom, "//total/vProd")
        .VL_SERV_NT = 0 ' Ajustar se houver campo específico
        .VL_TERC = 0
        .VL_DA = fnXML.ValidarValores(NFCom, "//total/vOutro")
        .VL_BC_ICMS = fnXML.ValidarValores(NFCom, "//total/ICMSTot/vBC")
        .VL_ICMS = fnXML.ValidarValores(NFCom, "//total/ICMSTot/vICMS")
        .COD_INF = ""
        .VL_PIS = fnXML.ValidarValores(NFCom, "//total/vPIS")
        .VL_COFINS = fnXML.ValidarValores(NFCom, "//total/vCOFINS")
        
        ' Campos Específicos Modelo 62
        .FIN_DOCE = fnXML.ValidarTag(NFCom, "//ide/finNFCom")
        .TIP_FAT = fnXML.ValidarTag(NFCom, "//ide/tpFat")
        
        ' Vínculos Pai (Hierarquia)
        .CHV_PAI_CONTRIBUICOES = assImportacao.ExtrairChaveRegD010()
        .CHV_PAI_FISCAL = assImportacao.ExtrairChaveRegD001()
        
        ' Tratamento de Cancelamento
        If VBA.Left(.COD_SIT, 2) = "02" Or VBA.Left(.COD_SIT, 2) = "03" Then
            .COD_PART = ""
            .VL_DOC = 0
            .VL_SERV = 0
            .VL_PIS = 0
            .VL_COFINS = 0
            .VL_BC_ICMS = 0
            .VL_ICMS = 0
        End If
        
        ' Gera a Chave Única do Registro D700
        .CHV_REG = fnSPED.GerarChaveRegistro(.ARQUIVO, "D700", .CHV_DOCE)
        
        ' Persistência no Dicionário Global
        If Not dtoRegSPED.rD700.Exists(.CHV_REG) Then
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_OPER, .IND_EMIT, _
                .COD_PART, .COD_MOD, .COD_SIT, "'" & .SER, "'" & .NUM_DOC, .DT_DOC, .DT_E_S, CDbl(.VL_DOC), _
                CDbl(.VL_DESC), CDbl(.VL_SERV), CDbl(.VL_SERV_NT), CDbl(.VL_TERC), CDbl(.VL_DA), CDbl(.VL_BC_ICMS), _
                CDbl(.VL_ICMS), .COD_INF, CDbl(.VL_PIS), CDbl(.VL_COFINS), .CHV_DOCE, .FIN_DOCE, .TIP_FAT, _
                .COD_MOD_DOC_REF, .CHV_DOCE_REF, .HASH_DOC_REF, .SER_DOC_REF, .NUM_DOC_REF, .MES_DOC_REF, .COD_MUN_DEST, CDbl(.DED))
                
            dtoRegSPED.rD700.Add .CHV_REG, Campos
        End If
        
        ' Processa Itens Filhos (Agregação para D730)
        ' Só processa itens se a nota for regular (não cancelada)
        If VBA.Left(.COD_SIT, 2) = "00" Then
            Call ProcessarItensD730(NFCom, .CHV_REG)
        End If
        
    End With
End Sub

Private Sub ProcessarItensD730(ByRef NFCom As DOMDocument60, ByVal ChavePai As String)
    ' Analítico dos Bilhetes (Agrupamento por CST, CFOP, Aliquota)
    
    Dim ListaDet As IXMLDOMNodeList
    Dim NoDet As IXMLDOMNode
    Dim DTO As clsAgregacaoD730
    Dim ChaveAgrupamento As String
    
    pDicionarioAgregacao.RemoveAll ' Limpa para a nota atual
    
    Set ListaDet = NFCom.SelectNodes("//det")
    
    ' 1. Loop de Agregação
    For Each NoDet In ListaDet
        ChaveAgrupamento = GerarChaveAgrupamento(NoDet)
        
        If pDicionarioAgregacao.Exists(ChaveAgrupamento) Then
            Set DTO = pDicionarioAgregacao(ChaveAgrupamento)
        Else
            Set DTO = New clsAgregacaoD730
            Dim Partes() As String
            Partes = VBA.Split(ChaveAgrupamento, "|")
            DTO.CST_ICMS = Partes(0)
            DTO.Cfop = Partes(1)
            DTO.ALIQ_ICMS = Partes(2)
            pDicionarioAgregacao.Add ChaveAgrupamento, DTO
        End If
        
        ' Acumula valores
        DTO.VL_OPR = DTO.VL_OPR + VBA.CDbl(fnXML.ValidarValores(NoDet, "vProd"))
        
        Dim NoImp As IXMLDOMNode: Set NoImp = NoDet.SelectSingleNode("imposto/ICMS00") ' Ajustar XPath se necessário
        If NoImp Is Nothing Then Set NoImp = NoDet.SelectSingleNode("imposto")
        
        If Not NoImp Is Nothing Then
            DTO.VL_BC_ICMS = DTO.VL_BC_ICMS + VBA.CDbl(fnXML.ValidarValores(NoImp, "vBC"))
            DTO.VL_ICMS = DTO.VL_ICMS + VBA.CDbl(fnXML.ValidarValores(NoImp, "vICMS"))
            DTO.VL_FCP = DTO.VL_FCP + VBA.CDbl(fnXML.ValidarValores(NoImp, "vFCP"))
        End If
    Next NoDet
    
    ' 2. Loop de Criação dos Registros
    Dim Chave As Variant
    For Each Chave In pDicionarioAgregacao.Keys
        Set DTO = pDicionarioAgregacao(Chave)
        Call CriarRegistroD730(DTO, ChavePai)
    Next Chave
    
End Sub

Private Sub CriarRegistroD730(ByVal DTO As clsAgregacaoD730, ByVal ChavePai As String)
    Dim Campos As Variant
    
    With CamposD730
        .REG = "D730"
        .ARQUIVO = DadosXML.ARQUIVO
        .CHV_PAI = ChavePai
        .CST_ICMS = DTO.CST_ICMS
        .Cfop = DTO.Cfop
        .ALIQ_ICMS = DTO.ALIQ_ICMS
        .VL_OPR = DTO.VL_OPR
        .VL_BC_ICMS = DTO.VL_BC_ICMS
        .VL_ICMS = DTO.VL_ICMS
        .VL_RED_BC = 0
        .COD_OBS = ""
        
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D730", .CST_ICMS, .Cfop, .ALIQ_ICMS)
        
        If Not dtoRegSPED.rD730.Exists(.CHV_REG) Then
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, .CHV_PAI_FISCAL, _
                           .CHV_PAI_CONTRIBUICOES, .CST_ICMS, .Cfop, CDbl(.ALIQ_ICMS), _
                           CDbl(.VL_OPR), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), CDbl(.VL_RED_BC), .COD_OBS)
            
            dtoRegSPED.rD730.Add .CHV_REG, Campos
        End If
        
        ' Se houver FCP, cria filho D731
        If DTO.VL_FCP > 0 Then
            Call CriarRegistroD731(.CHV_REG, DTO.VL_FCP)
        End If
    End With
End Sub

Private Sub CriarRegistroD731(ByVal ChavePaiD730 As String, ByVal ValorFCP As Double)
    Dim Campos As Variant
    
    With CamposD731
        .REG = "D731"
        .ARQUIVO = DadosXML.ARQUIVO
        .CHV_PAI = ChavePaiD730
        .VL_FCP_OP = ValorFCP
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D731")
        
        If Not dtoRegSPED.rD731.Exists(.CHV_REG) Then
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, CDbl(.VL_FCP_OP))
            dtoRegSPED.rD731.Add .CHV_REG, Campos
        End If
    End With
End Sub

' =========================================================================================
' 6. FUNÇÕES AUXILIARES
' =========================================================================================

Private Function GerarChaveAgrupamento(ByVal NoDet As IXMLDOMNode) As String
    Dim CST As String, Cfop As String, ALIQ As String
    Dim NoImposto As IXMLDOMNode, NoProd As IXMLDOMNode
    
    Set NoImposto = NoDet.SelectSingleNode("imposto")
    Set NoProd = NoDet.SelectSingleNode("prod")
    
    ' Tenta extrair CST de vários locais possíveis
    CST = fnXML.ValidarTag(NoDet, "imposto/*/CST")
    If CST = "" Then CST = "00" ' Default
    
    Cfop = fnXML.ValidarTag(NoProd, "CFOP")
    ALIQ = fnXML.ValidarValores(NoImposto, "*/pICMS")
    If ALIQ = "" Then ALIQ = "0"
    
    GerarChaveAgrupamento = CST & "|" & Cfop & "|" & ALIQ
End Function

Private Sub ExportarResultadosExcel()
    Dim Exportador As New ExportadorRegistros
    Call Exportador.ExportarRegistros("0000", "0000_Contr", "0001", "D001", "D700", "D730", "D731")
End Sub

Private Sub FinalizarProcesso()
    Call Util.MsgInformativa("Importação NFCom concluída com sucesso!", "Sucesso", VBA.Now)
    Call Util.AtualizarBarraStatus(False)
End Sub

