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
' ARQUITETURA: SLA (Single Level of Abstraction) com Orquestradores e Especialistas
' AUTOR: Equipe de Engenharia ControlDocs
' =========================================================================================

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private assImportacao As New AssistenteImportacao
Private GerenciadorSPED As New clsRegistrosSPED

' =========================================================================================
' 1. CAMADA PÚBLICA (ORQUESTRADOR NÍVEL 1)
' =========================================================================================

Public Sub ImportarNFCom(ByVal TipoImportacao As String)
    'On Error GoTo TratarErro

    Call CarregarDocumentos(TipoImportacao)
    Call CarregarDadosExistentes
    Call ExecutarProcessamentoEmLote
    Call ExportarResultadosExcel
    Call FinalizarProcesso

    Exit Sub
TratarErro:
    Call Util.MsgInformativa("Erro crítico na importação: " & VBA.Err.Description, "Erro", VBA.Now)
End Sub

Public Sub ProcessarNFCom(ByVal XML As DOMDocument60)
    'On Error GoTo TratarErroProcessamento

    Dim NoInfNFCom As IXMLDOMNode
    Dim ChavePai As String
    Dim DicionarioAgregacao As Scripting.Dictionary

    Set NoInfNFCom = ValidarEstrutura(XML)
    If NoInfNFCom Is Nothing Then Exit Sub

    Call PrepararAmbienteGlobal(XML)
    Call GarantirHierarquiaSPED(XML)
    Call GarantirParticipante(XML, NoInfNFCom)
    
    ChavePai = GerarRegistroD700(NoInfNFCom)
    
    Set DicionarioAgregacao = AgruparItensD730(NoInfNFCom)
    
    Call PersistirRegistrosFilhos(DicionarioAgregacao, ChavePai)

    Exit Sub
TratarErroProcessamento:
    Debug.Print "Erro ao processar arquivo (" & DadosXML.ARQUIVO & "): " & VBA.Err.Description
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

Private Sub CarregarDadosExistentes()
    
    Call Util.AtualizarBarraStatus("Carregando estrutura SPED existente...")
    
    With GerenciadorSPED
        
        Call .CarregarDadosRegistro0000("ARQUIVO")
        Call .CarregarDadosRegistro0000_Contr("ARQUIVO")
        
        Call .CarregarDadosRegistro0001("ARQUIVO")
        Call .CarregarDadosRegistroD001("ARQUIVO")
        Call .CarregarDadosRegistroD700("CHV_DOCE")
        
    End With
    
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
    Call Exportador.ExportarRegistros("0000", "0001", "0150", "0190", "0200", "D001", "D700", "D730", "D731")
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
    Dim DataEmissao As String
    Dim Ano As String, Mes As String
    
    ' 1. Limpeza inicial (Padrão DTO)
    Call DTO_DadosXML.ResetarDadosXML
    
    With DadosXML
        ' 2. Extração de Identificação (Tags específicas NFCom)
        .COD_MOD = fnXML.ValidarTag(XML, "//mod")
        .CNPJ_EMITENTE = fnXML.ValidarTag(XML, "//emit/CNPJ")
        .CNPJ_DESTINATARIO = fnXML.ValidarTag(XML, "//dest/CNPJ")
        If .CNPJ_DESTINATARIO = "" Then .CNPJ_DESTINATARIO = fnXML.ValidarTag(XML, "//dest/CPF")
        
        ' 3. Definição do Estabelecimento (Quem é o dono do arquivo?)
        ' Se o emitente for o contribuinte atual, é Saída Própria. Se não, é Entrada/Terceiros.
        .CNPJ_ESTABELECIMENTO = CNPJContribuinte ' Variável Global do Projeto
        
        ' 4. Extração de Data e Período (Fundamental para registro 0000)
        ' XPath NFCom: infNFCom/ide/dhEmi (Formato AAAA-MM-DD...)
        DataEmissao = VBA.Left(fnXML.ValidarTag(XML, "//ide/dhEmi"), 10)
        
        If VBA.Len(DataEmissao) >= 10 Then
            Ano = VBA.Left(DataEmissao, 4)
            Mes = VBA.Mid(DataEmissao, 6, 2)
            .Periodo = Mes & "/" & Ano
        Else
            ' Fallback se data vier vazia
            .Periodo = VBA.Format(Date, "mm/yyyy")
        End If
        
        ' 5. Definição do Nome do Arquivo Lógico (Chave para Dicionários 0000/0001)
        ' Padrão: MM/AAAA-CNPJ
        .ARQUIVO = .Periodo & "-" & .CNPJ_ESTABELECIMENTO
        
        ' 6. Tipo de Nota (Na NFCom não tem tpNF explícito como na NFe 0/1)
        ' Definimos baseado em quem emitiu:
        If .CNPJ_EMITENTE = .CNPJ_ESTABELECIMENTO Then
            .TIPO_NF = "1" ' Saída
        Else
            .TIPO_NF = "0" ' Entrada
        End If
        
    End With
    
End Sub

Private Sub GarantirHierarquiaSPED(ByVal XML As DOMDocument60)
    ' Verifica e cria Registro 0000 (Abertura do Arquivo)
    If Not dtoRegSPED.r0000.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0000(XML)
    End If
    
    If Not dtoRegSPED.r0000_Contr.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0000_Contr(XML)
    End If
    
    ' Verifica e cria Registro 0001 (Abertura do Bloco 0)
    If Not dtoRegSPED.r0001.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistro0001
    End If
    
    ' Verifica e cria Registro D001 (Abertura do Bloco D)
    If Not dtoRegSPED.rD001.Exists(DadosXML.ARQUIVO) Then
        Call assImportacao.CriarRegistroD001
    End If
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

Public Sub CriarRegistroD700(ByRef NFCom As DOMDocument60)

    Dim Campos As Variant
    Dim Chave As String
    
    ' Garante que o Dicionário esteja carregado antes de verificar existência
    If dtoRegSPED.rD700 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD700("CHV_DOCE")
    
    With CamposD700
        
        .REG = "D700"
        .ARQUIVO = DadosXML.ARQUIVO
        .COD_MOD = "62" ' NFCom
        .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(fnSPED.GerarCodigoSituacao(fnXML.ValidarSituacao(fnXML.ValidarTag(NFCom, "//cStat"))))
        .SER = VBA.Format(fnXML.ValidarTag(NFCom, "//ide/serie"), "000")
        .NUM_DOC = fnXML.ValidarTag(NFCom, "//ide/nNF")
        .DT_DOC = fnXML.ValidarTag(NFCom, "//ide/dhEmi")
        .DT_E_S = .DT_DOC ' Na NFCom geralmente são iguais
        
        ' Chave da Nota (Tag chNFCom ou ID)
        .CHV_DOCE = fnXML.ValidarTag(NFCom, "//chNFCom")
        If .CHV_DOCE = "" Then .CHV_DOCE = VBA.Replace(fnXML.ValidarTag(NFCom, "//infNFCom/@Id"), "NFCom", "")
        
        ' Participante
        .COD_PART = assImportacao.ExtrairCodigoParticipante()
        If .COD_PART <> "" Then .COD_PART = Util.FormatarTexto(.COD_PART)
        
        ' Valores Totais (XPath ajustado para NFCom)
        .VL_DOC = fnXML.ValidarValores(NFCom, "//total/vNF")
        .VL_DESC = fnXML.ValidarValores(NFCom, "//total/vDesc")
        .VL_SERV = fnXML.ValidarValores(NFCom, "//total/vProd")
        .VL_SERV_NT = 0 ' Geralmente calculado ou vOutro
        .VL_TERC = 0
        .VL_DA = fnXML.ValidarValores(NFCom, "//total/vOutro")
        .VL_BC_ICMS = fnXML.ValidarValores(NFCom, "//total/ICMSTot/vBC")
        .VL_ICMS = fnXML.ValidarValores(NFCom, "//total/ICMSTot/vICMS")
        .COD_INF = ""
        .VL_PIS = fnXML.ValidarValores(NFCom, "//total/vPIS")
        .VL_COFINS = fnXML.ValidarValores(NFCom, "//total/vCOFINS")
        
        ' Campos Específicos NFCom
        .FIN_DOCE = fnXML.ValidarTag(NFCom, "//ide/finNFCom")
        .TIP_FAT = fnXML.ValidarTag(NFCom, "//ide/tpFat")
        
        ' Chaves Pai
        .CHV_PAI_CONTRIBUICOES = assImportacao.ExtrairChaveRegD010() ' Ajustar se D010 não for o pai direto
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
        
        ' Geração da Chave do Registro
        .CHV_REG = fnSPED.GerarChaveRegistro(.ARQUIVO, "D700", .CHV_DOCE)
        
        ' Montagem do Array de Campos (Seguindo a ordem do D700)
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .IND_OPER, .IND_EMIT, _
            .COD_PART, .COD_MOD, .COD_SIT, "'" & .SER, "'" & .NUM_DOC, .DT_DOC, .DT_E_S, CDbl(.VL_DOC), _
            CDbl(.VL_DESC), CDbl(.VL_SERV), CDbl(.VL_SERV_NT), CDbl(.VL_TERC), CDbl(.VL_DA), CDbl(.VL_BC_ICMS), _
            CDbl(.VL_ICMS), .COD_INF, CDbl(.VL_PIS), CDbl(.VL_COFINS), .CHV_DOCE, .FIN_DOCE, .TIP_FAT, _
            .COD_MOD_DOC_REF, .CHV_DOCE_REF, .HASH_DOC_REF, .SER_DOC_REF, .NUM_DOC_REF, .MES_DOC_REF, .COD_MUN_DEST, CDbl(.DED))
            
        ' Persistência no Dicionário Global
        If Not dtoRegSPED.rD700.Exists(.CHV_DOCE) Then
            dtoRegSPED.rD700.Add .CHV_DOCE, Campos
        End If
        
    End With

End Sub

Private Function GerarRegistroD700(ByVal NoRaiz As IXMLDOMNode) As String
    Dim Registro As CamposRegD700
    Dim Chave As String
    
    With Registro
        .REG = "D700"
        .ARQUIVO = Util.NomeArquivo(DadosXML.ARQUIVO)
        .COD_MOD = "62"
        .COD_SIT = "00"
        .IND_OPER = "0"
        .IND_EMIT = "0"

        Call PreencherIdentificacaoD700(Registro, NoRaiz)
        Call PreencherValoresD700(Registro, NoRaiz)
        Call PreencherParticipanteD700(Registro, NoRaiz)
        
        Chave = fnXML.ValidarTag(NoRaiz, "chNFCom")
        If Chave = "" Then Chave = fnXML.ValidarTag(NoRaiz, "Id")
        .CHV_REG = fnSPED.GerarChaveRegistro(.ARQUIVO, "D700", Chave)
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
    Dados = Array(REG.REG, REG.ARQUIVO, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_FISCAL, _
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
        .ARQUIVO = Util.NomeArquivo(DadosXML.ARQUIVO)
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
    Dados = Array(REG.REG, REG.ARQUIVO, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_FISCAL, _
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
        .ARQUIVO = PaiD730.ARQUIVO
        .CHV_PAI = PaiD730.CHV_REG
        .VL_FCP_OP = VBA.Format(ValorFCP, "0.00")
        .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D731")
    End With
    
    Call SalvarD731NoGlobal(REG)
End Sub

Private Sub SalvarD731NoGlobal(ByRef REG As CamposRegD731)
    Dim Dados As Variant
    Dados = Array(REG.REG, REG.ARQUIVO, REG.CHV_REG, REG.CHV_PAI, REG.CHV_PAI_FISCAL, _
                  REG.CHV_PAI_CONTRIBUICOES, REG.VL_FCP_OP)

    If Not dtoRegSPED.rD731.Exists(REG.CHV_REG) Then
        dtoRegSPED.rD731.Add REG.CHV_REG, Dados
    End If
End Sub


