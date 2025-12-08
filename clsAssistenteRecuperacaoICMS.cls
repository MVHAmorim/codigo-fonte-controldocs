VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssistenteRecuperacaoICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dicTitulos As New Dictionary
Private arrRelatorio As New ArrayList

Public Function GerarRelatorioExclusaoICMS()

Dim Msg As String
Dim Inicio As Double
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    
    Call arrRelatorio.Clear
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    ' <<< MUDANÇA: O alvo agora é a nova planilha "Assistente Exclusão ICMS"
    Set dicTitulos = Util.MapearTitulos(assRecuperacaoICMS, 3)
    
    ' Carrega apenas os dados relevantes para a exclusão do ICMS Normal (C170)
    Call CarregarDadosC170_Normal(Msg)
    ' Futuramente, outras rotinas de carregamento (C181, C185, etc.) podem ser adicionadas aqui.
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        On Error Resume Next
            Dim shtRelatorio As Worksheet
            Set shtRelatorio = Sheets("Assistente Exclusão ICMS")
            If shtRelatorio.AutoFilter.FilterMode Then shtRelatorio.ShowAllData
        On Error GoTo 0
        
        Call Util.LimparDados(assRecuperacaoICMS, 4, False)
        Call Util.ExportarDadosArrayList(assRecuperacaoICMS, arrRelatorio)
        Call Util.MsgInformativa(arrRelatorio.Count & " registros elegíveis encontrados.", "Assistente de Exclusão do ICMS", Inicio)
        
    Else
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Verifique se o SPED foi importado e se existem operações elegíveis."
        Call Util.MsgAlerta(Msg, "Assistente de Exclusão do ICMS")
    End If
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Function

Private Sub CarregarDadosC170_Normal(ByVal Msg As String)

Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    
    b = 0
    With Apuracao
        
        Call .CarregarDadosRegistro0000(True)
        
        For Each Linha In Dados.Rows
            b = b + 1
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C170...", Dados.Rows.Count, Timer)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            
            ' Pula linhas em branco
            If Not Util.ChecarCamposPreenchidos(Campos) Then GoTo ProximaLinha
                
            ' Carrega variáveis para o filtro de elegibilidade
            Dim Cfop As String: Cfop = Campos(dicTitulosC170("CFOP"))
            Dim CST_ICMS As String: CST_ICMS = Campos(dicTitulosC170("CST_ICMS"))
            Dim CST_PIS As String: CST_PIS = Campos(dicTitulosC170("CST_PIS"))
            Dim VL_ICMS As Double: VL_ICMS = fnExcel.ConverterValores(Campos(dicTitulosC170("VL_ICMS")))
            Dim CHV_C100 As String: CHV_C100 = Campos(dicTitulosC170("CHV_PAI_CONTRIBUICOES"))
            Dim DT_ENT_SAI As String
            Dim DT_DOC As String
            
            Call .ExtrairDadosC100(CHV_C100)
            DT_ENT_SAI = .Campo(dicTitulos("DT_ENT_SAI"))
            DT_DOC = .Campo(dicTitulos("DT_DOC"))
            
            If DT_ENT_SAI <> "" Then DT_DOC = DT_ENT_SAI
            
            If isOperacaoElegivel_Normal(Cfop, CST_ICMS, CST_PIS, VL_ICMS, DT_DOC) Then
                
                ' Extrai todos os dados necessários (lógica similar à original)
                Dim REG As String: REG = Campos(dicTitulosC170("REG"))
                
                Dim Arquivo As String: Arquivo = Campos(dicTitulosC170("ARQUIVO"))
                Dim COD_ITEM As String: COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                Dim COD_PART As String: COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                Dim CNPJ_ESTABELECIMENTO As String: CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C100, Arquivo)
                Dim CHV_0140 As String: CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(Arquivo, CNPJ_ESTABELECIMENTO))
                
                Call .ExtrairDados0150_Contrib(CHV_0140, COD_PART, True)
                Call .ExtrairDados0200_Contrib(CHV_0140, COD_ITEM, True)
                
                ' Atribui valores aos campos do relatório
                .AtribuirValor "REG", Campos(dicTitulosC170("REG"))
                .AtribuirValor "ARQUIVO", Arquivo
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC170("CHV_REG"))
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "CFOP", Cfop
                .AtribuirValor "CST_ICMS", fnExcel.FormatarTexto(CST_ICMS)
                .AtribuirValor "VL_ITEM", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_ITEM")))
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_DESC")))
                
                ' <<< MUDANÇA: Adiciona o campo VL_ICMS ao relatório
                .AtribuirValor "VL_ICMS", VL_ICMS
                
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(CST_PIS)
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_PIS")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_COFINS")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_COFINS")))
                
                arrRelatorio.Add .Campo
            End If
ProximaLinha:
        Next Linha
    End With
    Set Apuracao = Nothing
End Sub

Private Function isOperacaoElegivel_Normal(ByVal Cfop As String, ByVal CST_ICMS As String, _
                                           ByVal CST_PIS As String, ByVal VL_ICMS As Double, _
                                           ByVal dtOperacao As String) As Boolean

Dim cstPisLimpo As String: cstPisLimpo = Trim(Util.ApenasNumeros(CST_PIS))
Dim dataCorte As Date
Dim dataOperacaoDate As Date
Dim dataValida As Boolean

' Data de corte da Lei 14.592/2023
dataCorte = CDate("2023-05-01")

' Converte a data da operação (String) para Date
dataValida = IsDate(dtOperacao)
If dataValida Then dataOperacaoDate = CDate(dtOperacao)

' Regra 1: Se não tem ICMS, não é elegível
If VL_ICMS <= 0 Then
    isOperacaoElegivel_Normal = False
    Exit Function
End If

' Regra 2: Define a elegibilidade com base no CST e na DATA
Select Case True
    
    ' Operações de Saída (Tese do Século)
    Case cstPisLimpo Like "01", cstPisLimpo Like "02"
        isOperacaoElegivel_Normal = True
        
    ' Operações de Entrada (Créditos - Lei 14.592/2023)
    Case cstPisLimpo Like "5*", cstPisLimpo Like "6*"
        ' Só é elegível para cálculo se a data for válida E
        ' a partir de 01/05/2023 (inclusive)
        If dataValida And (dataOperacaoDate >= dataCorte) Then
            isOperacaoElegivel_Normal = True
        Else
            ' Se a data for anterior ou inválida, não entra no relatório
            isOperacaoElegivel_Normal = False
        End If
        
    ' Outros casos
    Case Else
        isOperacaoElegivel_Normal = False
        
End Select
    
End Function

Sub CalcularExclusaoICMSNormal()

    Dim Dados As Range
    Dim Linha As Range
    Dim Campos As Variant
    Dim Inicio As Double
    Dim shtAlvo As Worksheet
    Dim dicTitulos As Dictionary
    Dim b As Long
    Dim linhasCalculadas As Long
    Dim linhasIgnoradas As Long
    
    Inicio = Now()
    Set shtAlvo = Sheets("Assistente Exclusão ICMS Normal")
    
    Call Util.DesabilitarControles
    Application.StatusBar = "Iniciando exclusão do ICMS da base de PIS/COFINS..."
    
    Set dicTitulos = Util.MapearTitulos(shtAlvo, 3)
    Set Dados = Util.DefinirIntervalo(shtAlvo, 4, 3)
    If Dados Is Nothing Then
        MsgBox "Nenhum dado encontrado para processar.", vbInformation
        GoTo Finalizar
    End If
    
    b = 0
    linhasCalculadas = 0
    linhasIgnoradas = 0
    
    For Each Linha In Dados.Rows
        b = b + 1
        Call Util.AntiTravamento(b, 100, "Calculando linha " & b & " de " & Dados.Rows.Count, Dados.Rows.Count, Inicio)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        
        Dim VlItem As Double: VlItem = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM")))
        Dim VlDesc As Double: VlDesc = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC")))
        Dim vlDesp As Double: vlDesp = fnExcel.ConverterValores(Campos(dicTitulos("VL_DESP")))
        Dim VlIcms As Double: VlIcms = fnExcel.ConverterValores(Campos(dicTitulos("VL_ICMS")))
        Dim aliqPis As Double: aliqPis = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_PIS")))
        Dim aliqCofins As Double: aliqCofins = fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_COFINS")))
        Dim vlPisOriginal As Double: vlPisOriginal = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS")))
        Dim vlCofinsOriginal As Double: vlCofinsOriginal = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS")))
        Dim basePisAtual As Double: basePisAtual = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_PIS")))
        Dim dtOperacao As String: dtOperacao = fnExcel.FormatarData(Campos(dicTitulos("DT_ENT_SAI")))
        Dim indOper As String: indOper = fnExcel.FormatarData(Campos(dicTitulos("IND_OPER")))
        
        Dim BaseOriginalOperacao As Double
        BaseOriginalOperacao = VlItem - VlDesc + vlDesp
        
        Dim NovaBaseCalculada As Double
        NovaBaseCalculada = BaseOriginalOperacao - VlIcms
        
        Dim NovaBase As Double
        NovaBase = IIf(NovaBaseCalculada < 0, 0, NovaBaseCalculada)
        
        If VBA.Round(NovaBase, 2) <> VBA.Round(basePisAtual, 2) Then
            Dim novoVlPis As Double: novoVlPis = VBA.Round(NovaBase * aliqPis, 2)
            Dim novoVlCofins As Double: novoVlCofins = VBA.Round(NovaBase * aliqCofins, 2)
            
            Campos(dicTitulos("VL_NOVA_BC_PIS")) = NovaBase
            Campos(dicTitulos("VL_NOVA_BC_COFINS")) = NovaBase
            Campos(dicTitulos("VL_NOVO_VL_PIS")) = novoVlPis
            Campos(dicTitulos("VL_NOVO_VL_COFINS")) = novoVlCofins
            Campos(dicTitulos("VL_PIS_RECUPERAR")) = vlPisOriginal - novoVlPis
            Campos(dicTitulos("VL_COFINS_RECUPERAR")) = vlCofinsOriginal - novoVlCofins
            Campos(dicTitulos("VL_TOTAL_RECUPERAR")) = (vlPisOriginal - novoVlPis) + (vlCofinsOriginal - novoVlCofins)
            
            linhasCalculadas = linhasCalculadas + 1
        Else
            ' Se as bases são iguais, limpa as colunas de resultado para não exibir dados redundantes.
            Campos(dicTitulos("VL_NOVA_BC_PIS")) = Empty
            Campos(dicTitulos("VL_NOVA_BC_COFINS")) = Empty
            Campos(dicTitulos("VL_NOVO_VL_PIS")) = Empty
            Campos(dicTitulos("VL_NOVO_VL_COFINS")) = Empty
            Campos(dicTitulos("VL_PIS_RECUPERAR")) = Empty
            Campos(dicTitulos("VL_COFINS_RECUPERAR")) = Empty
            Campos(dicTitulos("VL_TOTAL_RECUPERAR")) = Empty
            
            linhasIgnoradas = linhasIgnoradas + 1
        End If
        
        Linha.value = Campos
        
    Next Linha

    Dim msgFinal As String
    msgFinal = "Cálculo de exclusão concluído!" & vbCrLf & vbCrLf
    msgFinal = msgFinal & linhasCalculadas & " linhas tiveram a base recalculada." & vbCrLf
    msgFinal = msgFinal & linhasIgnoradas & " linhas já estavam com a base correta."
    
    Call Util.MsgInformativa(msgFinal, "Exclusão do ICMS Normal", Inicio)

Finalizar:
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles

End Sub

Public Sub AtualizarRegistros()

    Dim shtAlvo As Worksheet
    Dim dicTitulos As Dictionary
    Dim arrDados As New ArrayList
    Dim Campos As Variant
    Dim Inicio As Double
    Dim b As Long
    Dim registrosAtualizados As Long

    Inicio = Now()
    Set shtAlvo = Sheets("Assistente Exclusão ICMS Normal")

    ' --- Validação e Preparação ---
    Call Util.DesabilitarControles
    Application.StatusBar = "Iniciando a atualização modular dos registros do SPED..."

    ' Carrega em memória os registros do SPED que vamos manipular
    With New clsRegistrosSPED
        Call .CarregarDadosRegistroC100
        Call .CarregarDadosRegistroC170
        Call .CarregarDadosRegistroC175_Contr
        ' Adicione aqui o carregamento de outros registros quando precisar (ex: .CarregarDadosRegistroC181_Contr)
    End With

    Set dicTitulos = Util.MapearTitulos(shtAlvo, 3)
    Set arrDados = Util.CriarArrayListRegistro(shtAlvo)
    If arrDados.Count = 0 Then GoTo Finalizar

    registrosAtualizados = 0
    b = 0
    
    ' --- Loop Principal: Processa os dados do relatório ---
    For Each Campos In arrDados
        b = b + 1
        Call Util.AntiTravamento(b, 100, "Processando linha " & b & " de " & arrDados.Count, arrDados.Count, Inicio)

        ' Otimização: Só processa se houver valor a recuperar
        If fnExcel.ConverterValores(Campos(dicTitulos("VL_TOTAL_RECUPERAR"))) > 0 Then
            Dim REG As String: REG = CStr(Campos(dicTitulos("REG")))
            
            ' O Orquestrador decide qual especialista chamar
            Select Case REG
                Case "C170"
                    If AtualizarRegistroC170(Campos, dicTitulos) Then registrosAtualizados = registrosAtualizados + 1
                'Case "C175"
                    'If AtualizarRegistroC175(Campos, dicTitulos) Then registrosAtualizados = registrosAtualizados + 1
                ' --- PONTO DE EXPANSÃO FUTURA ---
                ' Case "C181"
                '   If AtualizarRegistroC181(Campos, dicTitulos) Then registrosAtualizados = registrosAtualizados + 1
            End Select
        End If
    Next Campos

    If registrosAtualizados > 0 Then
        ' --- Exporta os Dados Alterados de Volta para as Planilhas do SPED ---
        Dim ExpReg As New ExportadorRegistros
        Call ExpReg.ExportarRegistros("C170", "C175_Contr") ' Adicione outros REGs aqui
        Set ExpReg = Nothing
        
        ' --- Recalcula os Totais dos Documentos (C100) ---
        Application.StatusBar = "Recalculando totais dos documentos (Registro C100)..."
        Call rC170.AtualizarImpostosC100(True)
        'Call rC175Contr.AtualizarImpostosC100(True)
    End If
    
    ' --- Finalização ---
    Dim msgFinal As String
    msgFinal = registrosAtualizados & " registros do SPED foram atualizados com sucesso!"
    Call Util.MsgInformativa(msgFinal, "Atualização SPED Modular", Inicio)

Finalizar:
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles

End Sub

Private Function AtualizarRegistroC170(ByRef Campos As Variant, ByRef dicTitulos As Dictionary) As Boolean
    
    Dim CHV_REG As String: CHV_REG = CStr(Campos(dicTitulos("CHV_REG")))
    
    With dtoRegSPED
        ' Verifica se o registro C170 existe no SPED carregado em memória
        If .rC170.Exists(CHV_REG) Then
            Dim dicCamposC170 As Variant
            dicCamposC170 = .rC170(CHV_REG)
            
            ' Atualiza os campos com os novos valores calculados
            dicCamposC170(dtoTitSPED.tC170("VL_BC_PIS")) = Campos(dicTitulos("VL_NOVA_BC_PIS"))
            dicCamposC170(dtoTitSPED.tC170("VL_PIS")) = Campos(dicTitulos("VL_NOVO_VL_PIS"))
            dicCamposC170(dtoTitSPED.tC170("VL_BC_COFINS")) = Campos(dicTitulos("VL_NOVA_BC_COFINS"))
            dicCamposC170(dtoTitSPED.tC170("VL_COFINS")) = Campos(dicTitulos("VL_NOVO_VL_COFINS"))
            
            ' Devolve o array atualizado para o dicionário
            .rC170(CHV_REG) = dicCamposC170
            AtualizarRegistroC170 = True ' Sinaliza que houve atualização
        End If
    End With
    
End Function
