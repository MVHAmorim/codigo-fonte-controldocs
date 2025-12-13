VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportadorSPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================================
' CLASSE: clsImportadorSPED
' DESCRIÇÃO: Importador universal SPED (Fiscal/Contribuições).
' ALGORITMO: Leitura Stream (FSO) -> Processamento Global (fnSPED) -> Exportação Global.
' AUTOR: Engenharia ControlDocs (Refatoração Lógica)
' =========================================================================================

Public Enum EnumTipoSPED
    spFiscal = 1
    spContribuicoes = 2
End Enum

' Propriedades Privadas
Private pTipoSPED As EnumTipoSPED
Private pArquivosSelecionados As Collection
Private pPeriodo As String ' Armazena período (DT_INI a DT_FIN) capturado do 0000

' Dependências
Private FSO As Object

' =========================================================================================
' PROPRIEDADES
' =========================================================================================

Public Property Let Tipo(ByVal Valor As EnumTipoSPED)
    pTipoSPED = Valor
End Property

Public Property Get Tipo() As EnumTipoSPED
    Tipo = pTipoSPED
End Property

' =========================================================================================
' CONSTRUTOR / DESTRUTOR
' =========================================================================================

Private Sub Class_Initialize()
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    Set pArquivosSelecionados = New Collection
End Sub

Private Sub Class_Terminate()
    Set FSO = Nothing
    Set pArquivosSelecionados = Nothing
    ' Nota: dicRegistros é global, não anulamos aqui para permitir persistência se necessário,
    ' mas a limpeza é feita no início da importação.
End Sub

' =========================================================================================
' MÉTODO PÚBLICO
' =========================================================================================

Public Sub Importar(Optional ByVal Origem As String = "Arquivo")
    On Error GoTo TratarErroImportacao

    Call ValidarPreRequisitos
    Call SelecionarArquivos(Origem)
    
    If pArquivosSelecionados.Count = 0 Then Exit Sub
    
    Call CarregarEstruturas
    Call ProcessarListaArquivos
    Call ExportarResultados
    Call FinalizarProcesso
    
    Exit Sub
    
TratarErroImportacao:
    Call Util.MsgInformativa("Erro crítico na importação: " & VBA.Err.Description, "Erro", VBA.Now)
    Call Util.AtualizarBarraStatus(False)
End Sub

' =========================================================================================
' MÉTODOS PRIVADOS
' =========================================================================================

Private Sub ValidarPreRequisitos()
    ' Implementar validações de ambiente
End Sub

Private Sub SelecionarArquivos(ByVal Origem As String)
    Dim Caminho As Variant
    Set pArquivosSelecionados = New Collection
    
    Call ProcessarDocumentos.CarregarSPEDs(Origem)
    
    DocsFiscais.arrSPEDs.addRange DocsFiscais.arrSPEDFiscal
    DocsFiscais.arrSPEDs.addRange DocsFiscais.arrSPEDContribuicoes
    
    If DocsFiscais.arrSPEDs.Count > 0 Then
        For Each Caminho In DocsFiscais.arrSPEDs
            pArquivosSelecionados.Add Caminho
        Next Caminho
    End If
End Sub

Private Sub CarregarEstruturas()
    Call Util.AtualizarBarraStatus("Preparando ambiente global...")
    ' Limpa o dicionário GLOBAL de registros
    ' Assumindo que dicRegistros é Resetado ou Limpo aqui
    Set dicRegistros = New Dictionary
End Sub

Private Sub ProcessarListaArquivos()
    Dim ARQUIVO As Variant
    Dim Indice As Long
    Dim TempoInicio As Double
    
    Indice = 0
    TempoInicio = VBA.Timer
    
    For Each ARQUIVO In pArquivosSelecionados
        Call Util.AntiTravamento(Indice, 1, "Processando Arq " & (Indice + 1), pArquivosSelecionados.Count, TempoInicio)
        Call ProcessarArquivo(VBA.CStr(ARQUIVO))
        Indice = Indice + 1
    Next ARQUIVO
End Sub

Private Sub FinalizarProcesso()
    Call Util.AtualizarBarraStatus(False)
    Call Util.MsgInformativa("Processamento concluído com sucesso!", "Sucesso", VBA.Now)
End Sub

' =========================================================================================
' PROCESSAMENTO STREAM
' =========================================================================================

Private Sub ProcessarArquivo(ByVal CaminhoArquivo As String)
    Dim ArquivoTexto As Object
    Dim Linha As String
    Dim ContadorLinhas As Long
    
    On Error GoTo ErroArquivo
    
    Set ArquivoTexto = FSO.OpenTextFile(CaminhoArquivo, 1, False)
    ContadorLinhas = 0
    
    Do Until ArquivoTexto.AtEndOfStream
        Linha = ArquivoTexto.ReadLine
        ContadorLinhas = ContadorLinhas + 1
        
        If VBA.Len(Linha) > 2 And VBA.Left(Linha, 1) = "|" Then
            ' Captura Período do Registro 0000 para contexto
            If VBA.Mid(Linha, 2, 4) = "0000" Then CapturarPeriodo (Linha)
            
            Call ProcessarLinhaDinamica(Linha)
        End If
        
        If ContadorLinhas Mod 2000 = 0 Then
            Application.StatusBar = "Lendo linha " & ContadorLinhas & "..."
            DoEvents
        End If
    Loop
    
    ArquivoTexto.Close
    Set ArquivoTexto = Nothing
    Exit Sub

ErroArquivo:
    Debug.Print "Erro arquivo " & CaminhoArquivo & ": " & VBA.Err.Description
    If Not ArquivoTexto Is Nothing Then ArquivoTexto.Close
End Sub

Private Sub CapturarPeriodo(ByVal Linha As String)
    On Error Resume Next
    Dim Dados As Variant
    Dados = VBA.Split(Linha, "|")
    ' Formato SPED: |0000|COD_VER|COD_FIN|DT_INI|DT_FIN|...
    ' Indices: 0, 1="0000", 2, 3, 4=DT_INI, 5=DT_FIN
    If UBound(Dados) >= 5 Then
        pPeriodo = Dados(4) ' Pode-se concatenar se necessário: Dados(4) & "|" & Dados(5)
    End If
End Sub

Private Sub ProcessarLinhaDinamica(ByVal Linha As String)
    ' Delega processamento para o manipulador GLOBAL (fnSPED)
    ' Ele é quem sabe formatar datas, valores e chaves para o dicRegistros global
    Call fnSPED.ProcessarRegistro(Linha, pPeriodo, (pTipoSPED = spContribuicoes))
End Sub

' =========================================================================================
' EXPORTAÇÃO (VIA DICIONÁRIO GLOBAL)
' =========================================================================================

Private Sub ExportarResultados()
    Dim ChaveRegistro As Variant
    Dim PlanilhaDestino As Object
    
    Call Util.AtualizarBarraStatus("Exportando dados processados...")
    
    ' Itera sobre a variável GLOBAL dicRegistros (referenciada diretamente ou via FuncoesSPED?)
    ' O código do usuário menciona 'dicRegistros' que é Public em VariaveisGlobais
    
    For Each ChaveRegistro In dicRegistros.Keys
        If PlanilhaExiste(VBA.CStr(ChaveRegistro)) Then
            Set PlanilhaDestino = ThisWorkbook.Worksheets(VBA.CStr(ChaveRegistro))
            
            ' Exporta o ArrayList armazenado no dicionário global para a planilha
            Call Util.ExportarDadosArrayList(PlanilhaDestino, dicRegistros(ChaveRegistro))
            
            Debug.Print "Global Export: " & ChaveRegistro
        End If
    Next ChaveRegistro
End Sub

Private Function PlanilhaExiste(ByVal Nome As String) As Boolean
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(Nome)
    PlanilhaExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

