VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegrasTributariasICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ValidarCampo_ALIQ_ICMS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_ICMS As Double, ALIQ_ICMS_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, Cfop$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    ALIQ_ICMS = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_ICMS") - i))
    Cfop = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação do ICMS
    
    ALIQ_ICMS_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_ICMS") - i))
    
    Select Case True
        
        Case ALIQ_ICMS <> ALIQ_ICMS_TRIB
            INCONSISTENCIA = "ALIQ_ICMS divergente: " & fnExcel.FormatarPercentuais(ALIQ_ICMS) & " (informado) vs " & fnExcel.FormatarPercentuais(ALIQ_ICMS_TRIB) & " (cadastrado) para a operação com CFOP " & Cfop
            SUGESTAO = "Aplicar a alíquota do ICMS cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_CST_ICMS(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrCST_ICMS As New ArrayList
Dim CST_ICMS As String, CST_ICMS_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, Cfop$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    CST_ICMS = Campos(dicTitulosApuracao("CST_ICMS") - i)
    Cfop = Campos(dicTitulosApuracao("CFOP") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do ICMS
    CST_ICMS_TRIB = CamposTrib(dicTitulosTributacao("CST_ICMS") - t)
    
    Select Case True
        
        Case CST_ICMS <> CST_ICMS_TRIB
            INCONSISTENCIA = "CST_ICMS divergente: " & CST_ICMS & " (informado) vs " & CST_ICMS_TRIB & " (cadastrado) para a operação com CFOP " & Cfop
            SUGESTAO = "Aplicar o CST_ICMS cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ST_SEM_LIBERACAO(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)

Dim i As Byte, t As Byte
Dim arrST_SEM_LIBERACAO As New ArrayList
Dim ST_SEM_LIBERACAO As String, ST_SEM_LIBERACAO_TRIB$
Dim INCONSISTENCIA As String, SUGESTAO$, COD_NCM$
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    ST_SEM_LIBERACAO = Campos(dicTitulosApuracao("ST_SEM_LIBERACAO") - i)
    COD_NCM = Campos(dicTitulosApuracao("COD_NCM") - i)
    
    If LBound(CamposTrib) = 0 Then t = 1
    
    'Carregar campos do assistente de Tributação do ICMS
    ST_SEM_LIBERACAO_TRIB = CamposTrib(dicTitulosTributacao("ST_SEM_LIBERACAO") - t)
    
    Select Case True
        
        Case ST_SEM_LIBERACAO <> ST_SEM_LIBERACAO_TRIB
            INCONSISTENCIA = "ST_SEM_LIBERACAO divergente: " & ST_SEM_LIBERACAO & " (informado) vs " & ST_SEM_LIBERACAO_TRIB & " (cadastrado) para a operação com o NCM " & COD_NCM
            SUGESTAO = "Aplicar o ST_SEM_LIBERACAO cadastrado na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function

Public Function ValidarCampo_ALIQ_ST(ByRef Campos As Variant, ByRef dicTitulosApuracao As Dictionary, _
    ByRef CamposTrib As Variant, ByRef dicTitulosTributacao As Dictionary)
    
Dim ALIQ_ST As Double, ALIQ_ST_TRIB#
Dim INCONSISTENCIA As String, SUGESTAO$, Cfop$
Dim i As Byte
    
    If LBound(Campos) = 0 Then i = 1
    
    'Carregar campos do assistente de Apuração do ICMS
    ALIQ_ST = fnExcel.FormatarPercentuais(Campos(dicTitulosApuracao("ALIQ_ST") - i))
    Cfop = Campos(dicTitulosApuracao("CFOP") - i)

    If LBound(Campos) = 0 Then i = 1
    'Carregar campos do assistente de Tributação do ICMS
    
    ALIQ_ST_TRIB = fnExcel.FormatarPercentuais(CamposTrib(dicTitulosTributacao("ALIQ_ST") - i))
    
    Select Case True
        
        Case ALIQ_ST <> ALIQ_ST_TRIB
            INCONSISTENCIA = "ALIQ_ST divergente: " & fnExcel.FormatarPercentuais(fnExcel.FormatarPercentuais(ALIQ_ST)) & " (informado) vs " & fnExcel.FormatarPercentuais(fnExcel.FormatarPercentuais(ALIQ_ST_TRIB)) & " (cadastrado) para a operação com CFOP " & Cfop
            SUGESTAO = "Aplicar a alíquota do ICMS-ST cadastrada na Tributação"
            
    End Select
    
    If INCONSISTENCIA <> "" Then Call Util.GravarSugestao(Campos, dicTitulosApuracao, _
        INCONSISTENCIA:=INCONSISTENCIA, SUGESTAO:=SUGESTAO, dicInconsistenciasIgnoradas:=dicInconsistenciasIgnoradas)
        
End Function
