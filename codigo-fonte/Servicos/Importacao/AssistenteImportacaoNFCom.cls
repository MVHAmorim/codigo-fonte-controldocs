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
' DESCRIÇÃO: Importação e mapeamento de Notas Fiscais de Comunicação (Mod 21/22)
' PADRÃO: Orquestração Procedural (Matching NFe/CTe Architecture)
' =========================================================================================

Private EnumContribuicoes As New clsEnumeracoesSPEDContribuicoes
Private EnumFiscal As New clsEnumeracoesSPEDFiscal
Private GerenciadorSPED As New clsRegistrosSPED
Private assImportacao As New AssistenteImportacao
Private ExpReg As New ExportadorRegistros

' Variáveis de Contexto
Private CNPJBase As String
Private UF As String

' =========================================================================================
' 1. ORQUESTRADOR PRINCIPAL
' =========================================================================================

Public Sub ImportarNFCom(ByVal tpImportacao As String)
    'On Error GoTo TratarErro
    
    ' 1. Carregamento
    Call ProcessarDocumentos.CarregarXMLS(tpImportacao)
    If tpImportacao = "Arquivo" Then DocsFiscais.arrNFCom.addRange DocsFiscais.arrTodos
    
    If DocsFiscais.arrNFCom.Count = 0 Then Exit Sub
    
    CNPJBase = VBA.Left(CNPJContribuinte, 8)
    
    ' 2. Inicialização
    Call InicializarObjetos
    Call Util.AtualizarBarraStatus("Iniciando processamento de NFCom...")
    
    ' 3. Processamento
    Call ProcessarXMLS
    
    ' 4. Exportação
    ' Exporta D700 (Pai), D730 (Analítico), D731 (FCP)
    Call ExpReg.ExportarRegistros("0000", "0000_Contr", "D001", "D700", "D730", "D731", "D990")
    
    Call LimparObjetos
    Call Util.AtualizarBarraStatus("Importação de NFCom concluída.")
    
Exit Sub
'TratarErro:
    'Call GerenciadorErros.TratarErroPersonalizado("AssistenteImportacaoNFCom", "ImportarNFCom")
End Sub

Private Sub ProcessarXMLS()

Dim b As Long
Dim XML As Variant
Dim NFe As DOMDocument60
    
    b = 0
    Comeco = Timer
    DocsSemValidade = 0
    
    For Each XML In DocsFiscais.arrNFCom
        
        Set NFe = assImportacao.ExtrairDadosXML(XML)
        Call CriarRegistroD700(NFe)
        Call Util.AntiTravamento(b, 50, "Importando XML " & b + 1 & " de " & DocsFiscais.arrNFCom.Count, DocsFiscais.arrNFCom.Count, Comeco)
    
    Next XML
    
End Sub

Private Sub ProcessarXML(ByRef NFe As DOMDocument60)

Dim Chave As String
    
    If dtoRegSPED.rD700 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD700("IND_OPER", "IND_EMIT", "CHV_DOCE")
    
    With CamposD700
        
        ' Usamos os helpers do AssistenteImportacao para manter a lógica DRY (regras centralizadas)
        .IND_OPER = assImportacao.IdentificarTipoOperacao() ' 0-Entrada / 1-Saída
        .IND_EMIT = assImportacao.IdentificarTipoEmissao(NFe) ' 0-Própria / 1-Terceiros
        .CHV_DOCE = fnXML.ValidarChaveAcesso(NFe, "NFCom")
        
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_DOCE)
        .CHV_REG = Chave
        
        If dtoRegSPED.rD700.Exists(Chave) Then
            Call CarregarCamposD700(NFe, Chave)
        Else
            Call CriarRegistroD700(NFe)
        End If
        
        ' 5. Orquestração de Registros Filhos
        ' Se tiver participante identificado, cria o registro 0150 (Cadastro Participante)
        If .COD_PART <> "" Then Call assImportacao.CriarRegistro0150(NFe, .COD_PART)
        
        ' Processa Itens (Serviços/Produtos)
        Call CriarRegistroD730(NFe, Chave)
        
        ' Espaço para futuros registros filhos (Ex: D710, D720, etc.)
        
    End With
    
End Sub

Private Sub CarregarCamposD700(ByRef NFe As DOMDocument60, ByVal Chave As String)

Dim Campos As Variant
    
    ' Recupera o array de dados armazenado no Dicionário
    Campos = dtoRegSPED.rD700(Chave)
    
    With CamposD700
        
        .REG = Campos(dtoTitSPED.tD700("REG"))
        .ARQUIVO = Campos(dtoTitSPED.tD700("ARQUIVO"))
        .CHV_REG = Campos(dtoTitSPED.tD700("CHV_REG"))
        .CHV_PAI_FISCAL = Campos(dtoTitSPED.tD700("CHV_PAI_FISCAL"))
        .CHV_PAI_CONTRIBUICOES = Campos(dtoTitSPED.tD700("CHV_PAI_CONTRIBUICOES"))
        .IND_OPER = Campos(dtoTitSPED.tD700("IND_OPER"))
        .IND_EMIT = Campos(dtoTitSPED.tD700("IND_EMIT"))
        .COD_PART = Campos(dtoTitSPED.tD700("COD_PART"))
        .COD_MOD = Campos(dtoTitSPED.tD700("COD_MOD"))
        .COD_SIT = Campos(dtoTitSPED.tD700("COD_SIT"))
        .SER = Campos(dtoTitSPED.tD700("SER"))
        .NUM_DOC = Campos(dtoTitSPED.tD700("NUM_DOC"))
        .DT_DOC = Campos(dtoTitSPED.tD700("DT_DOC"))
        .DT_A_P = Campos(dtoTitSPED.tD700("DT_A_P"))
        .VL_DOC = CDbl(Campos(dtoTitSPED.tD700("VL_DOC")))
        .VL_DESC = CDbl(Campos(dtoTitSPED.tD700("VL_DESC")))
        .VL_SERV = CDbl(Campos(dtoTitSPED.tD700("VL_SERV")))
        .VL_SERV_NT = CDbl(Campos(dtoTitSPED.tD700("VL_SERV_NT")))
        .VL_TERC = CDbl(Campos(dtoTitSPED.tD700("VL_TERC")))
        .VL_DA = CDbl(Campos(dtoTitSPED.tD700("VL_DA")))
        .VL_BC_ICMS = CDbl(Campos(dtoTitSPED.tD700("VL_BC_ICMS")))
        .VL_ICMS = CDbl(Campos(dtoTitSPED.tD700("VL_ICMS")))
        .COD_INF = Campos(dtoTitSPED.tD700("COD_INF"))
        .VL_PIS = CDbl(Campos(dtoTitSPED.tD700("VL_PIS")))
        .VL_COFINS = CDbl(Campos(dtoTitSPED.tD700("VL_COFINS")))
        .CHV_DOCE = Campos(dtoTitSPED.tD700("CHV_DOCE"))
        .FIN_DOCE = Campos(dtoTitSPED.tD700("FIN_DOCE"))
        .TIP_FAT = Campos(dtoTitSPED.tD700("TIP_FAT"))
        .COD_MOD_DOC_REF = Campos(dtoTitSPED.tD700("COD_MOD_DOC_REF"))
        .CHV_DOCE_REF = Campos(dtoTitSPED.tD700("CHV_DOCE_REF"))
        .HASH_DOC_REF = Campos(dtoTitSPED.tD700("HASH_DOC_REF"))
        .SER_DOC_REF = Campos(dtoTitSPED.tD700("SER_DOC_REF"))
        .NUM_DOC_REF = Campos(dtoTitSPED.tD700("NUM_DOC_REF"))
        .MES_DOC_REF = Campos(dtoTitSPED.tD700("MES_DOC_REF"))
        .COD_MUN_DEST = Campos(dtoTitSPED.tD700("COD_MUN_DEST"))
        .DED = Campos(dtoTitSPED.tD700("DED"))
        
    End With
    
End Sub

Public Sub CriarRegistroD700(ByRef NFe As DOMDocument60)

Dim Campos As Variant
Dim Chave As String
Dim CNPJ_Part As String
    
    ' 1. Lazy Loading do DTO (Carrega definições se for a primeira execução)
    If dtoRegSPED.rD700 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD700("IND_OPER", "IND_EMIT", "CHV_DOCE")
    
    With CamposD700
        
        .REG = "D700"
        .ARQUIVO = DadosXML.ARQUIVO
        
        ' --- Identificação e Participantes ---
        ' Nota: IND_OPER e IND_EMIT já devem vir preenchidos do orquestrador (ProcessarXML)
        ' ou podem ser recalculados aqui para garantir isolamento.
        ' Mantendo a lógica do C100 de autossuficiência:
        .IND_OPER = assImportacao.IdentificarTipoOperacao()
        .IND_EMIT = assImportacao.IdentificarTipoEmissao(NFe)
        
        If .IND_EMIT = "0" Then
            ' Emissão Própria: Part é o Destinatário
            CNPJ_Part = fnXML.ValidarTag(NFe, "//dest/CNPJ")
            If CNPJ_Part = "" Then CNPJ_Part = fnXML.ValidarTag(NFe, "//dest/CPF")
        Else
            ' Terceiros: Part é o Emitente
            CNPJ_Part = fnXML.ValidarTag(NFe, "//emit/CNPJ")
        End If
        .COD_PART = Util.FormatarTexto(CNPJ_Part)
        
        .COD_MOD = fnXML.ValidarTag(NFe, "//infNFCom/ide/mod")
        
        ' Tratamento de Situação (cStat -> COD_SIT SPED)
        ' Exige mapeamento na EnumContribuicoes ou uso direto
        .COD_SIT = "00" ' Padrão Regular
        If fnXML.ValidarTag(NFe, "//cSit") <> "" Then
             ' Implementar fnSPED.GerarCodigoSituacao se necessário para NFCom
             ' .COD_SIT = EnumContribuicoes.ValidarEnumeracao_COD_SIT(...)
        End If
        
        .SER = VBA.Format(fnXML.ValidarTag(NFe, "//infNFCom/ide/serie"), "000")
        .NUM_DOC = fnXML.ValidarTag(NFe, "//infNFCom/ide/nNF")
        .DT_DOC = Util.FormatarData(fnXML.ValidarTag(NFe, "//infNFCom/ide/dhEmi"))
        .DT_A_P = .DT_DOC
        .VL_DOC = fnXML.ValidarValores(NFe, "//infNFCom/total/vNF")
        .VL_DESC = fnXML.ValidarValores(NFe, "//infNFCom/total/vDesc")
        .VL_SERV = fnXML.ValidarValores(NFe, "//infNFCom/total/vProd")
        .VL_SERV_NT = 0
        .VL_TERC = fnXML.ValidarValores(NFe, "//infNFCom/total/vOutro")
        .VL_DA = 0
        .VL_BC_ICMS = fnXML.ValidarValores(NFe, "//infNFCom/total/ICMSTot/vBC")
        .VL_ICMS = fnXML.ValidarValores(NFe, "//infNFCom/total/ICMSTot/vICMS")
        .COD_INF = ""
        .VL_PIS = fnXML.ValidarValores(NFe, "//infNFCom/total/vPIS")
        .VL_COFINS = fnXML.ValidarValores(NFe, "//infNFCom/total/vCOFINS")
        .FIN_DOCE = fnXML.ValidarTag(NFe, "//infNFCom/ide/tpEmis")
        .TIP_FAT = fnXML.ValidarTag(NFe, "//infNFCom/ide/tpFat")
        .COD_MUN_DEST = fnXML.ValidarTag(NFe, "//infNFCom/dest/enderDest/cMun")
        
        ' Referências (Placeholders inicializados)
        .COD_MOD_DOC_REF = ""
        .CHV_DOCE_REF = ""
        .HASH_DOC_REF = fnXML.ValidarTag(NFe, "//gRef/hash115")
        .SER_DOC_REF = ""
        .NUM_DOC_REF = ""
        .MES_DOC_REF = ""
        .DED = ""
        
        ' Hierarquia de Chaves
        .CHV_PAI_FISCAL = .CHV_DOCE ' Ou lógica de D001
        .CHV_PAI_CONTRIBUICOES = ""
        
        ' --- Lógica de Documento Cancelado (Cópia da lógica C100) ---
        ' Verifica se a chave está na lista global de cancelados
        If DocsFiscais.arrChavesCanceladas.contains(VBA.Replace(.CHV_DOCE, "'", "")) Then
            .COD_SIT = "02" ' Cancelado
        End If
        
        ' Se cancelado (02) ou inutilizado (03), zera valores fiscais
        If VBA.Left(.COD_SIT, 2) = "02" Or VBA.Left(.COD_SIT, 2) = "03" Then
            .COD_PART = ""
            .VL_DOC = 0
            .VL_DESC = 0
            .VL_SERV = 0
            .VL_SERV_NT = 0
            .VL_TERC = 0
            .VL_DA = 0
            .VL_BC_ICMS = 0
            .VL_ICMS = 0
            .VL_PIS = 0
            .VL_COFINS = 0
        End If
        
        Chave = Util.UnirCampos(.IND_OPER, .IND_EMIT, .CHV_DOCE)
        .CHV_REG = Chave
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                       .IND_OPER, .IND_EMIT, .COD_PART, .COD_MOD, .COD_SIT, "'" & .SER, "'" & .NUM_DOC, _
                       .DT_DOC, .DT_A_P, CDbl(.VL_DOC), CDbl(.VL_DESC), CDbl(.VL_SERV), CDbl(.VL_SERV_NT), _
                       CDbl(.VL_TERC), CDbl(.VL_DA), CDbl(.VL_BC_ICMS), CDbl(.VL_ICMS), .COD_INF, _
                       CDbl(.VL_PIS), CDbl(.VL_COFINS), .CHV_DOCE, .FIN_DOCE, .TIP_FAT, _
                       .COD_MOD_DOC_REF, .CHV_DOCE_REF, .HASH_DOC_REF, .SER_DOC_REF, _
                       .NUM_DOC_REF, .MES_DOC_REF, .COD_MUN_DEST, .DED)
                       
        ' Adiciona ao Dicionário Global (Sobrescreve se existir, C100 logic)
        dtoRegSPED.rD700(Chave) = Campos
        
    End With
    
End Sub

Private Sub CriarRegistroD730(ByRef NFe As DOMDocument60, ByVal ChavePai As String)
    ' D730: Registro Analítico (Agrupado por CST, CFOP, Aliq)
    ' D731: FCP (Filho do D730)
    
    Dim Detalhes As IXMLDOMNodeList
    Dim Det As IXMLDOMNode
    Dim dicAgregacao As Object ' Scripting.Dictionary
    Dim ChaveAgrup As String
    Dim CST As String, CFOP As String, ALIQ As String, COD_OBS As String
    Dim vOpr As Double, vBC As Double, vICMS As Double, vRedBC As Double, vFCP As Double
    Dim ItemAgregado As Variant
    Dim Key As Variant
    Dim Campos As Variant
    Dim ChaveD730 As String
    Dim i As Long
    
    Set dicAgregacao = CreateObject("Scripting.Dictionary")
    Set Detalhes = NFe.SelectNodes("//det")
    
    If dtoRegSPED.rD730 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD730("CHV_PAI_FISCAL", "CST_ICMS", "CFOP", "ALIQ_ICMS")
    If dtoRegSPED.rD731 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD731("CHV_PAI_FISCAL")
    
    ' 1. Agregação
    For Each Det In Detalhes
        CST = fnXML.ValidarTag(Det, "imposto/ICMS/*/CST")
        If CST = "" Then CST = "00" ' Fallback
        
        CFOP = fnXML.ValidarTag(Det, "prod/CFOP")
        ALIQ = fnXML.ValidarPercentual(Det, "imposto/ICMS//pICMS")
        If ALIQ = "" Then ALIQ = "0"
        
        COD_OBS = "" ' Implementar se houver observações fiscais
        
        ' Valores do Item
        vOpr = Util.TratarValor(fnXML.ValidarTag(Det, "prod/vProd"))
        vBC = Util.TratarValor(fnXML.ValidarTag(Det, "imposto/ICMS/*/vBC"))
        vICMS = Util.TratarValor(fnXML.ValidarTag(Det, "imposto/ICMS/*/vICMS"))
        vRedBC = 0 ' Calcular se necessário (vProd - vBC) se houver redução explícita
        vFCP = Util.TratarValor(fnXML.ValidarTag(Det, "imposto/ICMS/*/vFCP"))
        
        ChaveAgrup = CST & "|" & CFOP & "|" & ALIQ & "|" & COD_OBS
        
        If Not dicAgregacao.Exists(ChaveAgrup) Then
            ' Array: CST, CFOP, ALIQ, VL_OPR, VL_BC, VL_ICMS, VL_RED, COD_OBS, VL_FCP
            dicAgregacao.Add ChaveAgrup, Array(CST, CFOP, ALIQ, vOpr, vBC, vICMS, vRedBC, COD_OBS, vFCP)
        Else
            ItemAgregado = dicAgregacao(ChaveAgrup)
            ItemAgregado(3) = ItemAgregado(3) + vOpr
            ItemAgregado(4) = ItemAgregado(4) + vBC
            ItemAgregado(5) = ItemAgregado(5) + vICMS
            ItemAgregado(6) = ItemAgregado(6) + vRedBC
            ItemAgregado(8) = ItemAgregado(8) + vFCP ' Acumula FCP
            dicAgregacao(ChaveAgrup) = ItemAgregado
        End If
    Next Det
    
    ' 2. Criação dos Registros
    i = 0
    For Each Key In dicAgregacao.Keys
        ItemAgregado = dicAgregacao(Key)
        i = i + 1
        
        ' Gera chave única para o D730
        ChaveD730 = ChavePai & "|" & i
        
        With CamposD730
            .REG = "D730"
            .ARQUIVO = DadosXML.ARQUIVO
            .CHV_REG = ChaveD730
            .CHV_PAI_FISCAL = ChavePai
            .CHV_PAI_CONTRIBUICOES = ""
            
            .CST_ICMS = ItemAgregado(0)
            .CFOP = ItemAgregado(1)
            .ALIQ_ICMS = CDbl(ItemAgregado(2))
            .VL_OPR = CDbl(ItemAgregado(3))
            .VL_BC_ICMS = CDbl(ItemAgregado(4))
            .VL_ICMS = CDbl(ItemAgregado(5))
            .VL_RED_BC = CDbl(ItemAgregado(6))
            .COD_OBS = ItemAgregado(7)
            
            Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, _
                           .CST_ICMS, .CFOP, .ALIQ_ICMS, .VL_OPR, .VL_BC_ICMS, .VL_ICMS, .VL_RED_BC, .COD_OBS)
            
            dtoRegSPED.rD730.Add .CHV_REG, Campos
            
            ' 3. Processa Filho D731 (FCP) se houver valor
            vFCP = CDbl(ItemAgregado(8))
            If vFCP > 0 Then
                Call CriarRegistroD731(ChaveD730, vFCP)
            End If
        End With
    Next Key
    
    Set dicAgregacao = Nothing
End Sub

Private Sub CriarRegistroD731(ByVal ChavePaiD730 As String, ByVal ValorFCP As Double)
    Dim Campos As Variant
    Dim ChaveD731 As String
    
    ChaveD731 = ChavePaiD730 & "|FCP"
    
    With CamposD731
        .REG = "D731"
        .ARQUIVO = DadosXML.ARQUIVO
        .CHV_REG = ChaveD731
        .CHV_PAI_FISCAL = ChavePaiD730
        .CHV_PAI_CONTRIBUICOES = ""
        .VL_FCP_OP = ValorFCP
        
        Campos = Array(.REG, .ARQUIVO, .CHV_REG, .CHV_PAI_FISCAL, .CHV_PAI_CONTRIBUICOES, .VL_FCP_OP)
        
        dtoRegSPED.rD731.Add .CHV_REG, Campos
    End With
End Sub

' =========================================================================================
' 3. HELPER METHODS (Gestão de Objetos)
' =========================================================================================

Private Sub InicializarObjetos()
    ' Garante que o dicionário SPED está pronto para receber dados
    Call GerenciadorSPED.InicializarDicionarios
    Call GerenciadorSPED.CarregarTítulos("D700", "D730", "D731")
    ' Carrega DTOs específicos se ainda não existirem
    If dtoRegSPED.rD700 Is Nothing Then Call GerenciadorSPED.CarregarDadosRegistroD700("IND_OPER", "IND_EMIT", "CHV_DOCE")
End Sub

Private Sub LimparObjetos()
    Set NFe = Nothing
End Sub

