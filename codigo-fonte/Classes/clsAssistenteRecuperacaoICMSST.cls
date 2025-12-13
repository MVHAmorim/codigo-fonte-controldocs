VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssistenteRecuperacaoICMSST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dicTitulos As New Dictionary
Private arrRelatorio As New ArrayList

Public Function GerarRelatorioRecuperacaoICMSST()

Dim Msg As String
    
    Inicio = Now()
    
    Call Util.DesabilitarControles
    
    Call arrRelatorio.Clear
    Call dicInconsistenciasIgnoradas.RemoveAll
    'Call DadosValidacaoCFOP.CarregarTitulosRelatorio(ActiveSheet)
    Call DTO_RegistrosSPED.ResetarRegistrosSPED
    
    Set dicTitulos = Util.MapearTitulos(assRecuperacaoICMSST, 3)
    
    Call CarregarDadosC170(Msg)
    Call CarregarDadosC175(Msg)
    'Call CarregarDadosC181(Msg)
    'Call CarregarDadosC185(Msg)
    
    Application.StatusBar = "Processo concluído com sucesso!"
    If arrRelatorio.Count > 0 Then
        
        On Error Resume Next
            If assRecuperacaoICMSST.AutoFilter.FilterMode Then assRecuperacaoICMSST.ShowAllData
        On Error GoTo 0
        Call Util.LimparDados(assRecuperacaoICMSST, 4, False)
        
        Call Util.ExportarDadosArrayList(assRecuperacaoICMSST, arrRelatorio)
        Call FuncoesFormatacao.DestacarInconsistencias(assRecuperacaoICMSST)
        Call Util.MsgInformativa("Relatório gerado com sucesso", "Assistente de Recuperação do ICMS-ST", Inicio)
        
    Else
        
        Msg = "Nenhum dado encontrado para geração do relatório." & vbCrLf & vbCrLf
        Msg = Msg & "Por favor verifique se o SPED foi importado e tente novamente."
        Call Util.MsgAlerta(Msg, "Assistente de Recuperação do ICMS-ST")
        
    End If
    
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles
    
End Function

Private Sub CarregarDadosC170(ByVal Msg As String)

Dim ARQUIVO As String, REG$, CHV_REG$, CHV_0140$, CHV_0150$, CHV_C100$, COD_ITEM$, COD_PART$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, CST_ICMS$, Cfop$, CST_PIS$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC170 As New Dictionary
Dim Dados As Range, Linha As Range
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
    
    Set Dados = Util.DefinirIntervalo(regC170, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC170 = Util.MapearTitulos(regC170, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000(True)
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 10, Msg & "Carregando dados do registro C170", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                'Carrega as variáveis necessárias
                REG = Campos(dicTitulosC170("REG"))
                ARQUIVO = Campos(dicTitulosC170("ARQUIVO"))
                CHV_C100 = Campos(dicTitulosC170("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO(REG, CHV_C100, ARQUIVO)
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_0150 = Util.UnirCampos(CHV_0140, COD_PART)
                .UFContrib = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                COD_ITEM = Campos(dicTitulosC170("COD_ITEM"))
                CST_ICMS = Campos(dicTitulosC170("CST_ICMS"))
                Cfop = Campos(dicTitulosC170("CFOP"))
                CST_PIS = Campos(dicTitulosC170("CST_PIS"))
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC170("VL_ITEM")))
                'COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150_Contrib(CHV_0140, COD_PART, True)
                
                'Extrai dados do registro 0200
                Call .ExtrairDados0200_Contrib(CHV_0140, COD_ITEM, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "REG", REG
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC170("CHV_REG"))
                .AtribuirValor "COD_ITEM", fnExcel.FormatarTexto(COD_ITEM)
                .AtribuirValor "CFOP", Cfop
                .AtribuirValor "CST_ICMS", fnExcel.FormatarTexto(CST_ICMS)
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(CST_PIS)
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC170("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_PIS")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("ALIQ_COFINS")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC170("VL_COFINS")))
                
            End If
            
            If isOperacaoElegivel(Cfop, CST_ICMS, CST_PIS) Then arrRelatorio.Add .Campo
            
        Next Linha
        
    End With
    
    Set Apuracao = Nothing
    
End Sub

Private Sub CarregarDadosC175(ByVal Msg As String)

Dim ARQUIVO As String, CHV_REG$, CHV_PAI$, CHV_0140$, CHV_0150$, CHV_C100$, COD_ITEM$, COD_PART$, UF_CONTRIB$, COD_INC_TRIB$, CNPJ_ESTABELECIMENTO$, Cfop$, CST_PIS$
Dim Apuracao As New clsAssistenteApuracao
Dim dicTitulosC175 As New Dictionary
Dim Dados As Range, Linha As Range
Dim arrChavesPai As New ArrayList
Dim Comeco As Double, VL_ITEM#
Dim Campos As Variant
Dim b As Long
        
    Set Dados = Util.DefinirIntervalo(regC175_Contr, 4, 3)
    If Dados Is Nothing Then Exit Sub
    
    Set Apuracao.dicTitulos = dicTitulos
    Set dicTitulosC175 = Util.MapearTitulos(regC175_Contr, 3)
    
    b = 0
    Comeco = Timer
    With Apuracao
        
        Call .CarregarDadosRegistro0000
        Set arrChavesPai = Util.ListarValoresUnicos(regC170, 4, 3, "CHV_PAI_CONTRIBUICOES")
        
        For Each Linha In Dados.Rows
            
            .RedimensionarArray (dicTitulos.Count)
            Call Util.AntiTravamento(b, 50, Msg & "Carregando dados do registro C175", Dados.Rows.Count, Comeco)
            
            Campos = Application.index(Linha.Value2, 0, 0)
            If Util.ChecarCamposPreenchidos(Campos) Then
                
                CHV_PAI = Campos(dicTitulosC175("CHV_PAI_CONTRIBUICOES"))
                If arrChavesPai.contains(CHV_PAI) Then GoTo Prx: Stop
                
                'Carrega as variáveis necessárias
                ARQUIVO = Campos(dicTitulosC175("ARQUIVO"))
                UF_CONTRIB = .ExtrairUFContribuinte(ARQUIVO, True)
                CHV_C100 = Campos(dicTitulosC175("CHV_PAI_CONTRIBUICOES"))
                CNPJ_ESTABELECIMENTO = .ExtrairCNPJ_ESTABELECIMENTO_C100(CHV_C100)
                UF_CONTRIB = .ExtrairUF_Estabelecimento(ARQUIVO, CNPJ_ESTABELECIMENTO)
                COD_PART = .ExtrairCOD_PART_C100(CHV_C100)
                Cfop = Campos(dicTitulosC175("CFOP"))
                CST_PIS = Campos(dicTitulosC175("CST_PIS"))
                CHV_0140 = .ExtrairCHV_0140(Util.UnirCampos(ARQUIVO, CNPJ_ESTABELECIMENTO))
                CHV_0150 = Util.UnirCampos(CHV_0140, COD_PART)
                VL_ITEM = fnExcel.FormatarValores(Campos(dicTitulosC175("VL_OPER")))
                'COD_INC_TRIB = .ExtrairREGIME_TRIBUTARIO(ARQUIVO, True)
                
                'Extrai dados do registro C100
                Call .ExtrairDadosC100(CHV_C100)
                
                'Extrai dados do registro 0150
                Call .ExtrairDados0150_Contrib(CHV_0140, COD_PART, True)
                
                'Atribui valores aos campos do relatório
                .AtribuirValor "DESCR_ITEM", "O REGISTRO C175 NÃO POSSUI DADOS DO PRODUTO"
                .AtribuirValor "NOME_RAZAO", "O REGISTRO C175 NÃO POSSUI DADOS DO PARTICIPANTE"
                .AtribuirValor "REG", Campos(dicTitulosC175("REG"))
                .AtribuirValor "ARQUIVO", ARQUIVO
                .AtribuirValor "CHV_PAI_CONTRIBUICOES", CHV_C100
                .AtribuirValor "CHV_REG", Campos(dicTitulosC175("CHV_REG"))
                .AtribuirValor "CFOP", fnExcel.FormatarTexto(Cfop)
                .AtribuirValor "CST_ICMS", Empty
                .AtribuirValor "VL_ITEM", VL_ITEM
                .AtribuirValor "VL_DESP", .ExtrairVL_DESP_C100(CHV_C100, VL_ITEM)
                .AtribuirValor "VL_DESC", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_DESC")))
                .AtribuirValor "CST_PIS", fnExcel.FormatarTexto(CST_PIS)
                .AtribuirValor "CST_COFINS", fnExcel.FormatarTexto(Campos(dicTitulosC175("CST_COFINS")))
                .AtribuirValor "VL_BC_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_BC_PIS")))
                .AtribuirValor "ALIQ_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("ALIQ_PIS")))
                .AtribuirValor "VL_PIS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_PIS")))
                .AtribuirValor "VL_BC_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_BC_COFINS")))
                .AtribuirValor "ALIQ_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("ALIQ_COFINS")))
                .AtribuirValor "VL_COFINS", fnExcel.ConverterValores(Campos(dicTitulosC175("VL_COFINS")))
                
            End If
            
            If isOperacaoElegivel(Cfop, "", CST_PIS) Then arrRelatorio.Add .Campo
Prx:
        Next Linha
        
    End With
    
End Sub

Public Function isOperacaoElegivel(ByVal Cfop As String, ByVal CST_ICMS As String, ByVal CST_PIS As String) As Boolean

Dim cfopLimpo As String: cfopLimpo = Trim(Util.ApenasNumeros(Cfop))
Dim cstIcmsLimpo As String: cstIcmsLimpo = Trim(VBA.Right(Util.ApenasNumeros(CST_ICMS), 2))
Dim cstPisLimpo As String: cstPisLimpo = Trim(Util.ApenasNumeros(CST_PIS))

    ' --- Filtro 1: Validações de CFOP e CST PIS/COFINS (obrigatórias) ---
    Select Case cfopLimpo
        Case "5405", "5403", "5656", "6110", "6403", "6404"
            ' CFOP é elegível, continua a verificação
        Case Else
            isOperacaoElegivel = False ' CFOP inelegível, encerra a função
            Exit Function
    End Select
    
    Select Case cstPisLimpo
        Case "01", "02", "05"
            ' CST PIS/COFINS é elegível, continua a verificação
        Case Else
            isOperacaoElegivel = False ' CST PIS/COFINS inelegível, encerra a função
            Exit Function
    End Select
    
    ' --- Filtro 2: Validação Condicional do CST do ICMS ---
    If cstIcmsLimpo <> "" Then
        ' Se o CST_ICMS NÃO ESTÁ VAZIO, ele precisa ser 30 ou 60
        If cstIcmsLimpo = "30" Or cstIcmsLimpo = "60" Then
            isOperacaoElegivel = True ' Passou em todos os testes
        Else
            isOperacaoElegivel = False ' CST_ICMS preenchido, mas com valor inelegível
        End If
    Else
        ' Se o CST_ICMS ESTÁ VAZIO, a operação já é considerada elegível
        ' pois passou nos filtros obrigatórios de CFOP e CST PIS/COFINS.
        isOperacaoElegivel = True
    End If
    
End Function

Sub CalcularCreditoPresumidoICMSST()

    Dim shtAlvo As Worksheet
    Dim dicTitulos As Dictionary
    Dim rngAliquota As Range
    Dim Dados As Range
    Dim Linha As Range
    Dim Campos As Variant
    Dim Aliquota As Double
    Dim b As Long
    Dim Inicio As Double
    Dim linhasIgnoradas As Long, linhasCalculadas As Long
    
    Inicio = Now()
    Set shtAlvo = ActiveSheet
    
    ' --- Validações Iniciais (sem alteração) ---
    If shtAlvo.name <> "Assistente Exclusão ICMS-ST" Then
        MsgBox "Esta macro deve ser executada a partir da planilha 'Assistente Exclusão ICMS-ST'.", vbExclamation
        Exit Sub
    End If

    Call Util.DesabilitarControles
    Application.StatusBar = "Iniciando cálculo de crédito presumido..."

    ' --- ETAPA 1: Ler e Validar a Alíquota (sem alteração) ---
    On Error Resume Next
    Set rngAliquota = shtAlvo.Range("ALIQ_ICMS_ST")
    On Error GoTo 0

    If rngAliquota Is Nothing Then
        MsgBox "Erro: A célula nomeada 'ALIQ_ICMS_ST' não foi encontrada.", vbCritical
        GoTo Finalizar
    End If

    If IsEmpty(rngAliquota.value) Or Not IsNumeric(rngAliquota.value) Then
        MsgBox "Por favor, informe um valor numérico para a alíquota na célula nomeada 'ALIQ_ICMS_ST'.", vbExclamation
        rngAliquota.Select
        GoTo Finalizar
    End If
    
    Aliquota = CDbl(rngAliquota.value)
    If Aliquota <= 0 Or Aliquota >= 100 Then
        MsgBox "O valor da alíquota deve ser um número maior que 0 e menor que 100.", vbExclamation
        rngAliquota.Select
        GoTo Finalizar
    End If
    
    ' --- ETAPA 2: Preparar para o Processamento (sem alteração) ---
    Set dicTitulos = Util.MapearTitulos(shtAlvo, 3)
    Set Dados = Util.DefinirIntervalo(shtAlvo, 4, 3)
    If Dados Is Nothing Then
        MsgBox "Nenhum dado encontrado para processar.", vbInformation
        GoTo Finalizar
    End If
    
    b = 0
    linhasIgnoradas = 0
    linhasCalculadas = 0
    
    ' --- ETAPA 3: Loop OTIMIZADO com a LÓGICA FINAL ---
    For Each Linha In Dados.Rows
        b = b + 1
        Call Util.AntiTravamento(b, 100, "Verificando linha " & b & " de " & Dados.Rows.Count, Dados.Rows.Count, Inicio)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        
        ' --- INÍCIO DA LÓGICA DE VALIDAÇÃO E CÁLCULO ---
        
        ' 1. Define a base original da operação para o cálculo do ICMS
        Dim BaseOriginalOperacao As Double
        BaseOriginalOperacao = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM"))) - _
                             fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC"))) + _
                             fnExcel.ConverterValores(Campos(dicTitulos("VL_DESP")))
                             
        ' 2. Calcula qual DEVERIA SER o valor do ICMS Presumido com a alíquota atual
        Dim IcmsPresumidoCalculado As Double
        IcmsPresumidoCalculado = VBA.Round(BaseOriginalOperacao * (Aliquota), 2)
        
        ' 3. Calcula qual DEVERIA SER a base de PIS/COFINS após a exclusão
        Dim BaseCorretaAposExclusao As Double
        BaseCorretaAposExclusao = BaseOriginalOperacao - IcmsPresumidoCalculado
        
        ' 4. Validação: Compara a base que JÁ ESTÁ na planilha com a base que DEVERIA ESTAR
        Dim basePisAtual As Double
        basePisAtual = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_PIS")))
        
        If VBA.Round(basePisAtual, 2) = VBA.Round(BaseCorretaAposExclusao, 2) Then
            linhasIgnoradas = linhasIgnoradas + 1
            GoTo ProximaLinha ' Pula para o final do loop, pois a linha já está correta.
        End If

        ' --- EXECUÇÃO DO RECÁLCULO COMPLETO ---
        ' Se chegou até aqui, é porque a linha precisa ser (re)calculada.
        
        ' Preenche os campos do ICMS Presumido
        Campos(dicTitulos("ALIQ_ICMS_ST")) = Aliquota
        Campos(dicTitulos("VL_ICMS_ST_PRES")) = IcmsPresumidoCalculado
        
        ' Garante que a nova base não seja negativa
        Dim NovaBase As Double
        NovaBase = IIf(BaseCorretaAposExclusao < 0, 0, BaseCorretaAposExclusao)
        
        ' Atribui a nova base, calculada DO ZERO, para PIS e COFINS
        Campos(dicTitulos("VL_NOVA_BC_PIS")) = NovaBase
        Campos(dicTitulos("VL_NOVA_BC_COFINS")) = NovaBase
        
        ' Recalcula os valores de PIS e COFINS com base na NOVA base
        Campos(dicTitulos("VL_NOVO_VL_PIS")) = VBA.Round(NovaBase * (fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_PIS")))), 2)
        Campos(dicTitulos("VL_NOVO_VL_COFINS")) = VBA.Round(NovaBase * (fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_COFINS")))), 2)
        
        ' Calcula os valores a recuperar
        Campos(dicTitulos("VL_PIS_RECUPERAR")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS"))) - Campos(dicTitulos("VL_NOVO_VL_PIS"))
        Campos(dicTitulos("VL_COFINS_RECUPERAR")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS"))) - Campos(dicTitulos("VL_NOVO_VL_COFINS"))
        Campos(dicTitulos("VL_TOTAL_RECUPERAR")) = Campos(dicTitulos("VL_PIS_RECUPERAR")) + Campos(dicTitulos("VL_COFINS_RECUPERAR"))
        
        ' Atualiza a linha na planilha
        Linha.value = Campos
        linhasCalculadas = linhasCalculadas + 1
        
ProximaLinha:
    Next Linha

    Dim msgFinal As String
    msgFinal = "Cálculo de crédito presumido concluído!" & vbCrLf & vbCrLf
    msgFinal = msgFinal & linhasCalculadas & " linhas foram calculadas/recalculadas." & vbCrLf
    msgFinal = msgFinal & linhasIgnoradas & " linhas já estavam com os valores corretos e foram ignoradas."

    Call Util.MsgInformativa(msgFinal, "Exclusão do ICMS-ST", Inicio)

Finalizar:
    Call Util.AtualizarBarraStatus(False)
    Call Util.HabilitarControles

End Sub

Sub CalcularCreditoPresumidoICMSST_Old()

Dim shtAlvo As Worksheet
Dim dicTitulos As Dictionary
Dim rngAliquota As Range
Dim Dados As Range
Dim Linha As Range
Dim Campos As Variant
Dim Aliquota As Double
Dim b As Long
Dim Inicio As Double
    
    Inicio = Now()
    Set shtAlvo = ActiveSheet
    
    ' --- Validações Iniciais (perfeitas como estavam) ---
    If shtAlvo.name <> "Assistente Exclusão ICMS-ST" Then
        MsgBox "Esta macro deve ser executada a partir da planilha 'Assistente Exclusão ICMS-ST'.", vbExclamation
        Exit Sub
    End If

    Call Util.DesabilitarControles
    Application.StatusBar = "Iniciando cálculo de crédito presumido..."

    ' --- ETAPA 1: Ler e Validar a Alíquota (perfeito como estava) ---
    On Error Resume Next
    Set rngAliquota = shtAlvo.Range("ALIQ_ICMS_ST")
    On Error GoTo 0

    If rngAliquota Is Nothing Then
        MsgBox "Erro: A célula nomeada 'ALIQ_ICMS_ST' não foi encontrada.", vbCritical
        GoTo Finalizar
    End If

    If IsEmpty(rngAliquota.value) Or Not IsNumeric(rngAliquota.value) Then
        MsgBox "Por favor, informe um valor numérico para a alíquota na célula nomeada 'ALIQ_ICMS_ST'.", vbExclamation
        rngAliquota.Select
        GoTo Finalizar
    End If
    
    Aliquota = CDbl(rngAliquota.value)
    If Aliquota <= 0 Or Aliquota >= 100 Then
        MsgBox "O valor da alíquota deve ser um número maior que 0 e menor que 100.", vbExclamation
        rngAliquota.Select
        GoTo Finalizar
    End If
    
    ' --- ETAPA 2: Preparar para o Processamento ---
    Set dicTitulos = Util.MapearTitulos(shtAlvo, 3)
    Set Dados = Util.DefinirIntervalo(shtAlvo, 4, 3)
    If Dados Is Nothing Then
        MsgBox "Nenhum dado encontrado para processar.", vbInformation
        GoTo Finalizar
    End If
    
    b = 0
    ' --- ETAPA 3: Loop OTIMIZADO para Calcular TODAS as Linhas ---
    ' Como a planilha já está pré-filtrada, não precisamos mais do IF de elegibilidade aqui.
    ' Nós simplesmente calculamos para todas as linhas presentes no relatório.
    For Each Linha In Dados.Rows
        b = b + 1
        Call Util.AntiTravamento(b, 100, "Calculando linha " & b & " de " & Dados.Rows.Count, Dados.Rows.Count, Inicio)
        
        Campos = Application.index(Linha.Value2, 0, 0)
        
        ' --- EXECUÇÃO DIRETA DO CÁLCULO ---
        Dim BasePresumida As Double, IcmsPresumido As Double
        
        Campos(dicTitulos("ALIQ_ICMS_ST")) = Aliquota
        
        BasePresumida = fnExcel.ConverterValores(Campos(dicTitulos("VL_ITEM"))) - _
                        fnExcel.ConverterValores(Campos(dicTitulos("VL_DESC"))) + _
                        fnExcel.ConverterValores(Campos(dicTitulos("VL_DESP")))
                      
        IcmsPresumido = VBA.Round(BasePresumida * (Aliquota), 2)
        Campos(dicTitulos("VL_ICMS_ST_PRES")) = IcmsPresumido
        
        ' Recalcula bases e valores
        Dim VL_BC_PIS As Double: VL_BC_PIS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_PIS")))
        Dim VL_BC_COFINS As Double: VL_BC_COFINS = fnExcel.ConverterValores(Campos(dicTitulos("VL_BC_COFINS")))
        Campos(dicTitulos("VL_NOVA_BC_PIS")) = IIf(VL_BC_PIS - IcmsPresumido < 0, 0, VL_BC_PIS - IcmsPresumido)
        Campos(dicTitulos("VL_NOVA_BC_COFINS")) = IIf(VL_BC_COFINS - IcmsPresumido < 0, 0, VL_BC_COFINS - IcmsPresumido)
        Campos(dicTitulos("VL_NOVO_VL_PIS")) = VBA.Round(Campos(dicTitulos("VL_NOVA_BC_PIS")) * (fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_PIS")))), 2)
        Campos(dicTitulos("VL_NOVO_VL_COFINS")) = VBA.Round(Campos(dicTitulos("VL_NOVA_BC_COFINS")) * (fnExcel.ConverterValores(Campos(dicTitulos("ALIQ_COFINS")))), 2)
        Campos(dicTitulos("VL_PIS_RECUPERAR")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_PIS"))) - Campos(dicTitulos("VL_NOVO_VL_PIS"))
        Campos(dicTitulos("VL_COFINS_RECUPERAR")) = fnExcel.ConverterValores(Campos(dicTitulos("VL_COFINS"))) - Campos(dicTitulos("VL_NOVO_VL_COFINS"))
        Campos(dicTitulos("VL_TOTAL_RECUPERAR")) = Campos(dicTitulos("VL_PIS_RECUPERAR")) + Campos(dicTitulos("VL_COFINS_RECUPERAR"))
        
        ' Atualiza a linha na planilha com os novos dados calculados
        Linha.value = Campos
        
    Next Linha

    Call Util.MsgInformativa("Cálculo de crédito presumido concluído com sucesso!", "Exclusão do ICMS-ST", Inicio)

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
    Set shtAlvo = Sheets("Assistente Exclusão ICMS-ST")

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
                Case "C175"
                    If AtualizarRegistroC175(Campos, dicTitulos) Then registrosAtualizados = registrosAtualizados + 1
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
        Call rC175Contr.AtualizarImpostosC100(True)
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

Private Function AtualizarRegistroC175(ByRef Campos As Variant, ByRef dicTitulos As Dictionary) As Boolean

    Dim CHV_REG As String: CHV_REG = CStr(Campos(dicTitulos("CHV_REG")))
    
    With dtoRegSPED
        ' Verifica se o registro C175 existe no SPED carregado em memória
        If .rC175_Contr.Exists(CHV_REG) Then
            Dim dicCamposC175 As Variant
            dicCamposC175 = .rC175_Contr(CHV_REG)
            
            ' Atualiza os campos com os novos valores calculados
            dicCamposC175(dtoTitSPED.tC175_Contr("VL_BC_PIS")) = Campos(dicTitulos("VL_NOVA_BC_PIS"))
            dicCamposC175(dtoTitSPED.tC175_Contr("VL_PIS")) = Campos(dicTitulos("VL_NOVO_VL_PIS"))
            dicCamposC175(dtoTitSPED.tC175_Contr("VL_BC_COFINS")) = Campos(dicTitulos("VL_NOVA_BC_COFINS"))
            dicCamposC175(dtoTitSPED.tC175_Contr("VL_COFINS")) = Campos(dicTitulos("VL_NOVO_VL_COFINS"))
            
            ' Devolve o array atualizado para o dicionário
            .rC175_Contr(CHV_REG) = dicCamposC175
            AtualizarRegistroC175 = True ' Sinaliza que houve atualização
        End If
    End With

End Function


