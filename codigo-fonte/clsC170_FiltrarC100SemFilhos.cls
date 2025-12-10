VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsC170_FiltrarC100SemFilhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- Módulo para FILTRAR registros C100 sem C170 correspondentes ---
' Autor: Gemini
' Descrição: Este módulo fornece uma funcionalidade para analisar os dados
'            e FILTRAR na própria planilha C100 os registros
'            que não possuem nenhum registro filho C170 associado.

'--- Variáveis de Nível de Módulo ---
' Usadas para armazenar os dados carregados para a verificação.
Private dicTodosC100Keys As Object  ' Armazena as chaves de todos os registros C100
Private dicPaisC170 As Object       ' Armazena as chaves pai dos registros C170 para consulta rápida

'''
' Ponto de entrada principal. Orquestra o processo de filtragem.
'
Public Sub FiltrarC100SemFilhos()
    
    Dim Inicio As Date
    Dim arrChavesOrfas As Object
    Dim wsC100 As Worksheet
    
    Inicio = Now()
    Set wsC100 = ThisWorkbook.Worksheets("C100")
    
    ' Prepara o ambiente para execução otimizada
    Call Util.DesabilitarControles
    Call Util.AtualizarBarraStatus("Iniciando verificação de C100 sem C170...")
    
    ' 1. Carrega os dados necessários das planilhas para a memória
    If Not CarregarDados(wsC100) Then
        Call Util.HabilitarControles
        Call Util.AtualizarBarraStatus("Pronto")
        Exit Sub
    End If
    
    ' 2. Compara os dados e obtém uma lista de chaves C100 órfãs
    Set arrChavesOrfas = ListarChavesOrfas()
    
    ' 3. Aplica o filtro na planilha C100
    Call AplicarFiltro(wsC100, arrChavesOrfas)
    
    ' 4. Informa o usuário sobre o resultado
    If arrChavesOrfas.Count > 0 Then
        Call Util.MsgInformativa(arrChavesOrfas.Count & " registro(s) C100 sem filhos C170 foram filtrados com sucesso!", "Filtragem Concluída", Inicio)
    Else
        Call Util.MsgAviso("Nenhum registro C100 sem filhos C170 foi encontrado. Nenhum filtro foi aplicado.", "Verificação Concluída")
    End If
    
    ' Libera os recursos e restaura o ambiente do Excel
    Call LimparVariaveisGlobais
    Call Util.HabilitarControles
    Call Util.AtualizarBarraStatus("Pronto")
    
End Sub

'''
' Carrega as chaves dos registros C100 e C170 em dicionários para processamento rápido.
' Retorna: True se os dados foram carregados com sucesso, False caso contrário.
'
Private Function CarregarDados(ByRef wsC100 As Worksheet) As Boolean
    
    Dim wsC170 As Worksheet
    Dim Dados As Variant
    Dim dicTitulos As Object
    Dim i As Long
    Dim Chave As String
    
    On Error GoTo TratarErro
    
    ' Mapeia a planilha de origem C170
    Set wsC170 = ThisWorkbook.Worksheets("C170")
    
    ' Verifica se a planilha C100 contém dados
    If Util.PossuiDadosInformados(wsC100) = False Then
        Util.MsgAlerta "A planilha 'C100' não contém dados para análise.", "Dados Ausentes"
        CarregarDados = False
        Exit Function
    End If
    
    ' Carrega todas as chaves CHV_REG da C100
    Call Util.AtualizarBarraStatus("Carregando chaves C100...")
    Set dicTodosC100Keys = CreateObject("Scripting.Dictionary")
    Set dicTitulos = Util.MapearTitulos(wsC100, 3)
    Dados = Util.DefinirDados(wsC100, 4, 3)
    
    If IsArray(Dados) Then
        For i = LBound(Dados, 1) To UBound(Dados, 1)
            Chave = CStr(Dados(i, dicTitulos("CHV_REG")))
            If Not dicTodosC100Keys.Exists(Chave) Then
                dicTodosC100Keys.Add Chave, vbNullString
            End If
        Next i
    End If
    
    ' Carrega apenas as chaves pai (CHV_PAI_FISCAL) da C170 para uma busca rápida.
    Set dicPaisC170 = CreateObject("Scripting.Dictionary")
    If Util.PossuiDadosInformados(wsC170) Then
        Call Util.AtualizarBarraStatus("Mapeando filhos C170...")
        Set dicTitulos = Util.MapearTitulos(wsC170, 3)
        Dados = Util.DefinirDados(wsC170, 4, 3)
        
        If IsArray(Dados) Then
            For i = LBound(Dados, 1) To UBound(Dados, 1)
                Chave = CStr(Dados(i, dicTitulos("CHV_PAI_FISCAL")))
                If Not dicPaisC170.Exists(Chave) Then
                    dicPaisC170.Add Chave, vbNullString
                End If
            Next i
        End If
    End If
    
    CarregarDados = True
    Exit Function

TratarErro:
    Util.MsgCritica "Erro ao carregar dados." & vbCrLf & _
                   "Verifique se as planilhas 'C100' e 'C170' existem e estão nomeadas corretamente." & vbCrLf & _
                   "Erro: " & Err.Description, "Falha no Carregamento"
    CarregarDados = False
    
End Function

'''
' Compara os dicionários e retorna uma lista com as chaves dos C100 órfãos.
'
Private Function ListarChavesOrfas() As Object

    Dim arrChaves As New ArrayList
    Dim chaveC100 As Variant
    
    If dicTodosC100Keys Is Nothing Or dicTodosC100Keys.Count = 0 Then
        Set ListarChavesOrfas = arrChaves
        Exit Function
    End If
    
    Call Util.AtualizarBarraStatus("Identificando registros órfãos...")
    
    ' Percorre cada chave de C100
    For Each chaveC100 In dicTodosC100Keys.Keys
        ' Se a chave do C100 NÃO existe no dicionário de pais do C170, é um órfão.
        If Not dicPaisC170.Exists(chaveC100) Then
            arrChaves.Add chaveC100
        End If
    Next chaveC100
    
    Set ListarChavesOrfas = arrChaves
    
End Function

'''
' Aplica o autofiltro em uma planilha com base em uma lista de chaves.
'
Private Sub AplicarFiltro(ByRef Plan As Worksheet, ByRef arrChaves As Object)

    Dim dicTitulos As Object
    Dim rngFiltro As Range
    
    ' Garante que a planilha está ativa e limpa filtros anteriores
    Plan.Activate
    If Plan.AutoFilterMode Then Plan.AutoFilter.ShowAllData
    
    ' Se não houver chaves para filtrar, não faz nada
    If arrChaves.Count = 0 Then Exit Sub
    
    ' Define o intervalo de dados a ser filtrado (a partir da linha de títulos)
    Set rngFiltro = Util.DefinirIntervalo(Plan, 3, 3)
    If rngFiltro Is Nothing Then Exit Sub
    
    ' Mapeia os títulos para encontrar a coluna correta
    Set dicTitulos = Util.MapearTitulos(Plan, 3)
    
    ' Aplica o filtro na coluna "CHV_REG"
    rngFiltro.AutoFilter Field:=dicTitulos("CHV_REG"), Criteria1:=arrChaves.ToArray(), Operator:=xlFilterValues
    
End Sub


'''
' Libera a memória, limpando as variáveis de nível de módulo.
'
Private Sub LimparVariaveisGlobais()
    Set dicTodosC100Keys = Nothing
    Set dicPaisC170 = Nothing
End Sub



