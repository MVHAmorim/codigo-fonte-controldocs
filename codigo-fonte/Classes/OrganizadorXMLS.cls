VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrganizadorXMLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' === Constantes de configuração ===
Private Const NOME_PASTA_PERIODOS As String = "Periodos"

' === Ponto de entrada único ===
Public Sub RealocarArquivosXmlPorData()
    Dim pastaOrigem As String
    pastaOrigem = SelecionarPastaOrigem()
    If pastaOrigem = "" Then Exit Sub   ' usuário cancelou

    Dim pastaPeriodos As String
    pastaPeriodos = pastaOrigem & "\" & NOME_PASTA_PERIODOS
    CriarPastaSeNecessario pastaPeriodos

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    
    ' Percorre pasta e subpastas recursivamente
    PercorrerPasta FSO.GetFolder(pastaOrigem), pastaPeriodos
    
    MsgBox "Processo concluído!", vbInformation
End Sub

' === Processa um único arquivo XML ===
Private Sub ProcessarArquivo(ByVal ARQUIVO As Scripting.file, ByVal pastaPeriodos As String)
    On Error GoTo TratarErro
    
    Dim DataEmissao As Date
    DataEmissao = ExtrairDataDhEmi(ARQUIVO.Path)
    
    Dim pastaDestino As String
    pastaDestino = ObterPastaDestino(DataEmissao, pastaPeriodos)
    
    CriarPastaSeNecessario pastaDestino
    
    Dim caminhoDestino As String
    caminhoDestino = pastaDestino & "\" & ARQUIVO.name
    
    ' Copia apenas se ainda não existir no destino
    If Dir(caminhoDestino) = "" Then
        ARQUIVO.Copy caminhoDestino
    End If
    Exit Sub
    
TratarErro:
    Debug.Print "Erro ao processar " & ARQUIVO.name & ": " & Err.Description
End Sub

' === Extrai a data da tag <dhEmi> ===
Private Function ExtrairDataDhEmi(ByVal caminhoXml As String) As Date
    Dim xmlDoc As MSXML2.DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument60
    
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.Load caminhoXml
    
    Dim nodo As MSXML2.IXMLDOMNode
    Set nodo = xmlDoc.SelectSingleNode("//*[local-name()='dhEmi']")
    
    If nodo Is Nothing Then Err.Raise vbObjectError + 1, , "Tag <dhEmi> não encontrada."
    
    ' ISO 8601 › “2024-05-10T14:45:00-03:00” ? pegamos só a parte da data
    Dim parteData As String: parteData = Split(nodo.text, "T")(0)
    ExtrairDataDhEmi = CDate(parteData)
End Function

' === Retorna o caminho da subpasta destino no formato "mm-aaaa" ===
Private Function ObterPastaDestino(ByVal DataEmissao As Date, ByVal pastaPeriodos As String) As String
    Dim nomePasta As String
    nomePasta = Format(DataEmissao, "mm-yyyy")
    ObterPastaDestino = pastaPeriodos & "\" & nomePasta
End Function

' === Cria a pasta se ela não existir (DRY) ===
Private Sub CriarPastaSeNecessario(ByVal caminhoPasta As String)
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    
    If Not FSO.FolderExists(caminhoPasta) Then
        FSO.CreateFolder caminhoPasta
    End If
End Sub

' === Seleciona a pasta de origem ===
Private Function SelecionarPastaOrigem() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Selecione a pasta de origem dos XMLs"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelecionarPastaOrigem = .SelectedItems(1)
        Else
            SelecionarPastaOrigem = ""
        End If
    End With
End Function

' === Percorre pastas e subpastas recursivamente ===
Private Sub PercorrerPasta(ByVal pastaAtual As Scripting.Folder, ByVal pastaPeriodos As String)
    Dim ARQUIVO As Scripting.file
    For Each ARQUIVO In pastaAtual.Files
        If LCase(pastaAtual.Path) Like LCase("*" & NOME_PASTA_PERIODOS & "*") Then
            ' Evita processar a pasta Periodos
        ElseIf LCase(Right(ARQUIVO.name, 4)) = ".xml" Then
            ProcessarArquivo ARQUIVO, pastaPeriodos
        End If
    Next ARQUIVO
    
    Dim subpasta As Scripting.Folder
    For Each subpasta In pastaAtual.SubFolders
        PercorrerPasta subpasta, pastaPeriodos
    Next subpasta
End Sub

