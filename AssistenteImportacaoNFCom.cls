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

' Instância para reuso de lógica (ex: cadastro de participante 0150)
Private assImportacao As New AssistenteImportacao

' Estrutura para agregação de itens do registro D730 em memória
Private Type TAgregacaoD730
    ChaveAgrupamento As String
    CST_ICMS As String
    CFOP As String
    ALIQ_ICMS As String
    VL_OPR As Double
    VL_BC_ICMS As Double
    VL_ICMS As Double
    VL_FCP As Double
End Type

' Método principal para processar o XML da NFCom
Public Sub ProcessarNFCom(ByVal XML As DOMDocument60)
    On Error GoTo TratarErro
    
    Dim noInfNFCom As IXMLDOMNode
    Dim noDestinatario As IXMLDOMNode
    Dim noTotal As IXMLDOMNode
    Dim listaNosDetalhe As IXMLDOMNodeList
    Dim noDetalhe As IXMLDOMNode
    Dim noImposto As IXMLDOMNode
    Dim noProduto As IXMLDOMNode
    
    Dim RegistroD700 As CamposRegD700
    Dim RegistroD730 As CamposRegD730
    Dim RegistroD731 As CamposRegD731
    
    Dim CNPJ_Destinatario As String
    Dim CPF_Destinatario As String
    Dim ChaveAcesso As String
    
    Dim dicAgregacao As New Scripting.Dictionary
    Dim Chave As String
    Dim DadosItem As TAgregacaoD730
    Dim ItemAgregado As TAgregacaoD730
    
    ' Namespace handling - usando local-name para evitar problemas com prefixos
    Set noInfNFCom = XML.selectSingleNode("//*[local-name()='infNFCom']")
    
    If noInfNFCom Is Nothing Then
        Exit Sub
    End If
    
    ' ============================================================================================
    ' BLOCO D700 - CABEÇALHO
    ' ============================================================================================
    
    ' 1. Processar Participante (Destinatário)
    Set noDestinatario = noInfNFCom.selectSingleNode("*[local-name()='dest']")
    If Not noDestinatario Is Nothing Then
        CNPJ_Destinatario = fnXML.ValidarTag(noDestinatario, "CNPJ")
        CPF_Destinatario = fnXML.ValidarTag(noDestinatario, "CPF")
        
        If CNPJ_Destinatario <> "" Then
            ' Garante que o participante exista no registro 0150
            Call assImportacao.CriarRegistro0150(XML, CNPJ_Destinatario)
            RegistroD700.COD_PART = CNPJ_Destinatario
        ElseIf CPF_Destinatario <> "" Then
            Call assImportacao.CriarRegistro0150(XML, CPF_Destinatario)
            RegistroD700.COD_PART = CPF_Destinatario
        End If
    End If
    
    ' 2. Preencher Campos do D700
    With RegistroD700
        .REG = "D700"
        .Arquivo = Util.NomeArquivo(DadosXML.Arquivo)
        
        ' Identificadores Básicos
        ChaveAcesso = fnXML.ValidarTag(noInfNFCom, "chNFCom")
        ' Fallback para ID se a chave não estiver na tag direta (embora chNFCom seja o padrão)
        If ChaveAcesso = "" Then ChaveAcesso = fnXML.ValidarTag(noInfNFCom, "Id")
        
        .COD_MOD = "62"
        .SER = fnXML.ValidarTag(noInfNFCom, "serie")
        .NUM_DOC = fnXML.ValidarTag(noInfNFCom, "nNF")
        .DT_DOC = fnXML.ValidarTag(noInfNFCom, "dhEmi") ' Formato AAAA-MM-DD...
        .DT_E_S = .DT_DOC
        
        ' Valores Totais do Cabeçalho
        Set noTotal = noInfNFCom.selectSingleNode("*[local-name()='total']")
        If Not noTotal Is Nothing Then
            .VL_DOC = fnXML.ValidarValores(noTotal, "vNF")
            .VL_PIS = fnXML.ValidarValores(noTotal, "vPIS")
            .VL_COFINS = fnXML.ValidarValores(noTotal, "vCOFINS")
            .VL_SERV = fnXML.ValidarValores(noTotal, "vProd") ' Valor dos itens/serviços
        End If
        
        ' Situação do Documento
        ' 00 - Documento regular
        .COD_SIT = "00"
        
        ' Indicador de Operação: 0-Entrada (Aquisição)
        .IND_OPER = "0"
        
        ' Indicador do Emitente: 0-Emissão Própria (se emitente = declarante), mas na importação geralmente é Terceiros.
        ' Se for ImportadorXML, assumimos que estamos importando nota de TERCEIROS, então IND_EMIT = 1 (Terceiros).
        ' Porém, NFCom modelo 62 pode ser emissão própria.
        ' Vou manter "0" conforme padrão de entrada/saída direta ou lógica do projeto. 
        ' Se é importação para o SPED do PRÓPRIO emitente (saída), é 0.
        .IND_EMIT = "0" 
        
        ' Geração da Chave Única do Registro (SPED)
        ' Utiliza o helper fnSPED para consistência
        .CHV_REG = fnSPED.GerarChaveRegistro(.Arquivo, "D700", ChaveAcesso)
        
    End With
    
    ' Persistir D700 no Dicionário Global
    Dim arrayD700 As Variant
    arrayD700 = Array( _
        RegistroD700.REG, RegistroD700.Arquivo, RegistroD700.CHV_REG, RegistroD700.CHV_PAI, RegistroD700.CHV_PAI_Fiscal, _
        RegistroD700.CHV_PAI_CONTRIBUICOES, RegistroD700.IND_OPER, RegistroD700.IND_EMIT, RegistroD700.COD_PART, _
        RegistroD700.COD_MOD, RegistroD700.COD_SIT, RegistroD700.SER, RegistroD700.NUM_DOC, RegistroD700.DT_DOC, _
        RegistroD700.DT_E_S, RegistroD700.VL_DOC, RegistroD700.VL_DESC, RegistroD700.VL_SERV, RegistroD700.VL_SERV_NT, _
        RegistroD700.VL_TERC, RegistroD700.VL_DA, RegistroD700.VL_BC_ICMS, RegistroD700.VL_ICMS, RegistroD700.COD_INF, _
        RegistroD700.VL_PIS, RegistroD700.VL_COFINS, RegistroD700.CHV_DOCE, RegistroD700.FIN_DOCE, RegistroD700.TIP_FAT, _
        RegistroD700.COD_MOD_DOC_REF, RegistroD700.CHV_DOCE_REF, RegistroD700.HASH_DOC_REF, RegistroD700.SER_DOC_REF, _
        RegistroD700.NUM_DOC_REF, RegistroD700.MES_DOC_REF, RegistroD700.COD_MUN_DEST, RegistroD700.DED _
    )
    
    If Not dtoRegSPED.rD700.Exists(RegistroD700.CHV_REG) Then
        dtoRegSPED.rD700.Add RegistroD700.CHV_REG, arrayD700
    End If
    
    ' ============================================================================================
    ' BLOCO D730 - AGREGAÇÃO DE ITENS
    ' ============================================================================================
    
    Set listaNosDetalhe = noInfNFCom.selectNodes("*[local-name()='det']")
    
    For Each noDetalhe In listaNosDetalhe
        Set noImposto = noDetalhe.selectSingleNode("*[local-name()='imposto']")
        Set noProduto = noDetalhe.selectSingleNode("*[local-name()='prod']")
        
        Dim CST As String
        Dim CFOP_Item As String
        Dim Aliquota As String
        
        Dim valorProduto As Double
        Dim valorBC As Double
        Dim valorICMS As Double
        Dim valorFCP As Double
        
        ' 1. Extração de Dados do Item
        ' Tenta buscar CST no imposto (ICMSxx) ou direto no det se existir tag global
        CST = fnXML.ValidarTag(noDetalhe, "CST")
        If CST = "" And Not noImposto Is Nothing Then CST = fnXML.ValidarTag(noImposto, "CST")
        
        ' CFOP
        CFOP_Item = fnXML.ValidarTag(noDetalhe, "CFOP")
        If CFOP_Item = "" And Not noProduto Is Nothing Then CFOP_Item = fnXML.ValidarTag(noProduto, "CFOP")
        
        ' Valores
        valorProduto = CDbl(Val(fnXML.ValidarValores(noDetalhe, "vProd")))
        
        ' Tributos (ICMS)
        If Not noImposto Is Nothing Then
            valorBC = CDbl(Val(fnXML.ValidarValores(noImposto, "vBC")))
            valorICMS = CDbl(Val(fnXML.ValidarValores(noImposto, "vICMS")))
            valorFCP = CDbl(Val(fnXML.ValidarValores(noImposto, "vFCP")))
            Aliquota = fnXML.ValidarValores(noImposto, "pICMS")
        End If
        
        ' 2. Definição da Chave de Agrupamento
        Chave = CST & "|" & CFOP_Item & "|" & Aliquota
        
        ' 3. Acumular no Dicionário
        If dicAgregacao.Exists(Chave) Then
            ItemAgregado = dicAgregacao(Chave)
            
            ItemAgregado.VL_OPR = ItemAgregado.VL_OPR + valorProduto
            ItemAgregado.VL_BC_ICMS = ItemAgregado.VL_BC_ICMS + valorBC
            ItemAgregado.VL_ICMS = ItemAgregado.VL_ICMS + valorICMS
            ItemAgregado.VL_FCP = ItemAgregado.VL_FCP + valorFCP
            
            dicAgregacao(Chave) = ItemAgregado
        Else
            With DadosItem
                .ChaveAgrupamento = Chave
                .CST_ICMS = CST
                .CFOP = CFOP_Item
                .ALIQ_ICMS = Aliquota
                .VL_OPR = valorProduto
                .VL_BC_ICMS = valorBC
                .VL_ICMS = valorICMS
                .VL_FCP = valorFCP
            End With
            dicAgregacao.Add Chave, DadosItem
        End If
        
    Next noDetalhe
    
    ' ============================================================================================
    ' PERSISTÊNCIA D730 E D731 (MEMÓRIA PARA GLOBAL)
    ' ============================================================================================
    
    Dim chaveDicionario As Variant
    For Each chaveDicionario In dicAgregacao.Keys
        ItemAgregado = dicAgregacao(chaveDicionario)
        
        ' Preencher Registro D730
        With RegistroD730
            .REG = "D730"
            .Arquivo = RegistroD700.Arquivo
            .CHV_PAI = RegistroD700.CHV_REG
            
            .CST_ICMS = ItemAgregado.CST_ICMS
            .CFOP = ItemAgregado.CFOP
            .ALIQ_ICMS = ItemAgregado.ALIQ_ICMS
            
            .VL_OPR = Format(ItemAgregado.VL_OPR, "0.00")
            .VL_BC_ICMS = Format(ItemAgregado.VL_BC_ICMS, "0.00")
            .VL_ICMS = Format(ItemAgregado.VL_ICMS, "0.00")
            ' VL_RED_BC não foi calculado explicitamente, mantido zerado ou deve ser derivado
            .VL_RED_BC = "0,00" 
            
            ' Gerar Chave Única D730: Pai + Chaves de Agrupamento
            .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D730", .CST_ICMS, .CFOP, .ALIQ_ICMS)
        End With
        
        ' Array D730 para Dicionário Global
        Dim arrayD730 As Variant
        arrayD730 = Array( _
            RegistroD730.REG, RegistroD730.Arquivo, RegistroD730.CHV_REG, RegistroD730.CHV_PAI, RegistroD730.CHV_PAI_Fiscal, _
            RegistroD730.CHV_PAI_CONTRIBUICOES, RegistroD730.CST_ICMS, RegistroD730.CFOP, RegistroD730.ALIQ_ICMS, _
            RegistroD730.VL_OPR, RegistroD730.VL_BC_ICMS, RegistroD730.VL_ICMS, RegistroD730.VL_RED_BC, RegistroD730.COD_OBS _
        )
        
        If Not dtoRegSPED.rD730.Exists(RegistroD730.CHV_REG) Then
            dtoRegSPED.rD730.Add RegistroD730.CHV_REG, arrayD730
        End If
        
        ' Verificar necessidade do Registro D731 (FCP)
        If ItemAgregado.VL_FCP > 0 Then
            With RegistroD731
                .REG = "D731"
                .Arquivo = RegistroD730.Arquivo
                .CHV_PAI = RegistroD730.CHV_REG
                .VL_FCP_OP = Format(ItemAgregado.VL_FCP, "0.00")
                
                ' Chave D731
                .CHV_REG = fnSPED.GerarChaveRegistro(.CHV_PAI, "D731")
            End With
            
            Dim arrayD731 As Variant
            arrayD731 = Array( _
                RegistroD731.REG, RegistroD731.Arquivo, RegistroD731.CHV_REG, RegistroD731.CHV_PAI, _
                RegistroD731.CHV_PAI_Fiscal, RegistroD731.CHV_PAI_CONTRIBUICOES, RegistroD731.VL_FCP_OP _
            )
            
            If Not dtoRegSPED.rD731.Exists(RegistroD731.CHV_REG) Then
                dtoRegSPED.rD731.Add RegistroD731.CHV_REG, arrayD731
            End If
        End If
        
    Next chaveDicionario
    
    Exit Sub

TratarErro:
    ' Aqui poderia ter um Log de Erros
    Debug.Print "Erro em AssistenteImportacaoNFCom.ProcessarNFCom: " & Err.Description
End Sub
