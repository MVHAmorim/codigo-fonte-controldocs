Attribute VB_Name = "RecursosDesenvolvedor"
Option Explicit

' =========================================================================================
' MÓDULO: RecursosDesenvolvedor
' DESCRIÇÃO: Ferramentas DevOps para controle de versão do VBA e Dados.
' FUNCIONALIDADES:
'   - Exportação Inteligente (@Folder support)
'   - Importação Limpa e Segura (Ignora Planilhas/Documentos)
'   - Versionamento de Dados (.csv dos UsedRange)
'   - Otimização de Memória
' REQUISITOS:
'   - Microsoft Visual Basic for Applications Extensibility 5.3 (VBIDE)
'   - Microsoft Scripting Runtime (FileSystemObject)
' =========================================================================================

' Constantes VBIDE (Late Binding Safe)
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_ActiveXDesigner As Long = 11
Private Const vbext_ct_Document As Long = 100

Public Sub ExportarCodigoControlDocs(control As IRibbonControl)
    Dim vbProj As Object
    Dim VBComp As Object
    Dim FSO As Object
    Dim CaminhoBase As String, pastaDestino As String
    Dim extension As String
    Dim Contador As Long
    
    On Error GoTo TratarErro
    
    Set vbProj = Application.VBE.ActiveVBProject
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    CaminhoBase = "E:\Projetos\ControlDocs\codigo-fonte"
    
    Call CriarPastaSeNaoExistir(FSO, CaminhoBase)
    Application.StatusBar = "Exportando código fonte..."
    
    Contador = 0
    For Each VBComp In vbProj.VBComponents
        ' Define destino baseados no tipo
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                extension = ".cls"
                pastaDestino = CaminhoBase & "\Classes"
            Case vbext_ct_StdModule
                extension = ".bas"
                pastaDestino = CaminhoBase & "\Modulos"
            Case vbext_ct_MSForm
                extension = ".frm"
                pastaDestino = CaminhoBase & "\Formularios"
            Case vbext_ct_Document
                extension = ".cls"
                pastaDestino = CaminhoBase & "\Planilhas" ' Backup apenas (ReadOnly)
            Case Else
                extension = ".txt"
                pastaDestino = CaminhoBase & "\Outros"
        End Select
        
        ' Override com @Folder (exceto para Documentos, que forçamos em Planilhas por segurança)
        If VBComp.Type <> vbext_ct_Document Then
            Dim PastaUser As String
            PastaUser = LerAnotacaoFolder(VBComp)
            If VBA.Len(PastaUser) > 0 Then pastaDestino = CaminhoBase & "\" & PastaUser
        End If
        
        Call CriarPastaRecursiva(FSO, pastaDestino)
        
        ' Exporta
        ' Call Util.AntiTravamento -> Substituído por DoEvents para isolamento
        If Contador Mod 5 = 0 Then VBA.DoEvents
        
        VBComp.Export pastaDestino & "\" & VBComp.name & extension
        
        Contador = Contador + 1
        Set VBComp = Nothing
    Next VBComp
    
    ' Nova etapa: Exportar Dados das Planilhas
    Call ExportarDadosDasPlanilhas
    
    Application.StatusBar = False
    MsgBox "Exportação completa (Código e Dados)!" & vbCrLf & Contador & " arquivos exportados.", vbInformation, "Sucesso"
    
Limpar:
    Set VBComp = Nothing
    Set vbProj = Nothing
    Set FSO = Nothing
    Exit Sub
TratarErro:
    Debug.Print "Erro Exportar: " & VBA.Err.Description
    Resume Next
End Sub

Public Sub ImportarCodigoControlDocs(control As IRibbonControl)
    Dim vbProj As Object
    Dim VBComp As Object
    Dim FSO As Object
    Dim CaminhoBase As String
    Dim Contador As Long
    
    If MsgBox("ATENÇÃO: Isso apagará TODOS os módulos/classes (exceto Planilhas e este módulo). Continuar?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    
    On Error GoTo TratarErroImp
    
    Set vbProj = Application.VBE.ActiveVBProject
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    CaminhoBase = "E:\Projetos\ControlDocs\codigo-fonte"
    
    If Not FSO.FolderExists(CaminhoBase) Then
        MsgBox "Pasta codigo-fonte não encontrada.", vbCritical
        Exit Sub
    End If
    
    ' 1. Limpeza Segura (Clean Build)
    Application.StatusBar = "Limpando componentes antigos..."
    For Each VBComp In vbProj.VBComponents
        ' PROTEÇÃO CRÍTICA: Ignora Documentos (Planilhas/Workbook) e o próprio Módulo de Ferramentas
        If VBComp.Type = vbext_ct_Document Or VBComp.name = "RecursosDesenvolvedor" Then
            ' Não faz nada, mantém intacto
        Else
            vbProj.VBComponents.Remove VBComp
        End If
        Set VBComp = Nothing
    Next VBComp
    
    ' 2. Importação Recursiva (Ignorando pasta Planilhas)
    Application.StatusBar = "Importando novos componentes..."
    Contador = 0
    Call ImportarRecursivo(FSO.GetFolder(CaminhoBase), vbProj, Contador)
    
    Application.StatusBar = False
    MsgBox "Importação concluída! " & Contador & " arquivos importados.", vbInformation, "Sucesso"
    
    Exit Sub
TratarErroImp:
    MsgBox "Erro Importar: " & VBA.Err.Description, vbCritical
End Sub

Public Sub ExportarDadosDasPlanilhas()
    Dim ws As Object ' Worksheet (Late binding para evitar erro se referência Excel faltar, embora raro)
    Dim FSO As Object
    Dim CaminhoDados As String
    Dim NomeArquivo As String
    Dim ArquivoNum As Integer
    Dim LinhaStr As String
    Dim R As Long, c As Long
    Dim UltimaLinha As Long, ultimaColuna As Long
    Dim Valor As Variant
    
    On Error GoTo ErroDados
    
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    CaminhoDados = "E:\Projetos\ControlDocs\codigo-fonte\Dados"
    Call CriarPastaRecursiva(FSO, CaminhoDados)
    
    Application.StatusBar = "Exportando dados das planilhas..."
    
    For Each ws In ThisWorkbook.Worksheets
        NomeArquivo = SanitizarNome(ws.name) & ".csv"
        ArquivoNum = VBA.FreeFile
        
        Open CaminhoDados & "\" & NomeArquivo For Output As #ArquivoNum
        
        ' Usar UsedRange para exportar
        With ws.UsedRange
            UltimaLinha = .Rows.Count
            ultimaColuna = .Columns.Count
            
            For R = 1 To UltimaLinha
                LinhaStr = ""
                For c = 1 To ultimaColuna
                    Valor = .Cells(R, c).Value2
                    ' Tratamento básico de CSV (aspas e ponto e vírgula)
                    If VBA.VarType(Valor) = vbString Then
                        Valor = VBA.Replace(Valor, ";", ",") ' Evita quebra do CSV simples
                        Valor = VBA.Replace(Valor, vbCrLf, " ") ' Remove quebras de linha
                    End If
                    
                    LinhaStr = LinhaStr & VBA.CStr(Valor) & ";"
                Next c
                Print #ArquivoNum, LinhaStr
            Next R
        End With
        
        Close #ArquivoNum
    Next ws
    
    Exit Sub
ErroDados:
    Debug.Print "Erro ao exportar dados da planilha " & ws.name & ": " & VBA.Err.Description
    If ArquivoNum > 0 Then Close #ArquivoNum
    Resume Next
End Sub

' =========================================================================================
' HELPER FUNCTIONS (PRIVADAS E AUTOSSUFICIENTES)
' =========================================================================================

Private Sub ImportarRecursivo(ByVal Folder As Object, ByVal vbProj As Object, ByRef Count As Long)
    Dim file As Object, SubFolder As Object
    Dim Ext As String
    
    ' Pula pasta de backup de planilhas para evitar conflitos de importação
    If Folder.name = "Planilhas" Then Exit Sub
    
    For Each file In Folder.Files
        Ext = VBA.LCase(VBA.Right(file.name, 4))
        If Ext = ".cls" Or Ext = ".bas" Or Ext = ".frm" Then
            ' Não importa o próprio módulo para evitar duplicação/erro
            If file.name <> "RecursosDesenvolvedor.bas" Then
                ' Verifica se componente já existe (no caso de documentos) para não tentar sobrescrever incorretamente
                If Not ComponenteExiste(vbProj, VBA.Left(file.name, VBA.Len(file.name) - 4)) Then
                    On Error Resume Next
                    vbProj.VBComponents.Import file.Path
                    If VBA.Err.Number = 0 Then Count = Count + 1
                    On Error GoTo 0
                End If
            End If
        End If
    Next file
    
    For Each SubFolder In Folder.SubFolders
        Call ImportarRecursivo(SubFolder, vbProj, Count)
    Next SubFolder
End Sub

Private Function ComponenteExiste(ByVal vbProj As Object, ByVal Nome As String) As Boolean
    On Error Resume Next
    Dim TESTE As Object
    Set TESTE = vbProj.VBComponents(Nome)
    ComponenteExiste = Not TESTE Is Nothing
End Function

Private Function LerAnotacaoFolder(ByVal VBComp As Object) As String
    ' Lê as primeiras 5 linhas em busca de '@Folder("Caminho")
    Dim i As Long, Line As String, StartP As Long, EndP As Long
    On Error Resume Next
    If VBComp.CodeModule.CountOfLines = 0 Then Exit Function
    For i = 1 To 5
        If i > VBComp.CodeModule.CountOfLines Then Exit For
        Line = VBComp.CodeModule.Lines(i, 1)
        If VBA.InStr(1, Line, "@Folder(", vbTextCompare) > 0 Then
            StartP = VBA.InStr(1, Line, """") + 1
            EndP = VBA.InStr(StartP, Line, """")
            If EndP > StartP Then
                LerAnotacaoFolder = VBA.Mid(Line, StartP, EndP - StartP)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function SanitizarNome(ByVal Nome As String) As String
    Dim Invalidos As Variant, i As Long
    Invalidos = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    SanitizarNome = Nome
    For i = LBound(Invalidos) To UBound(Invalidos)
        SanitizarNome = VBA.Replace(SanitizarNome, Invalidos(i), "_")
    Next i
End Function

Private Sub CriarPastaSeNaoExistir(ByVal FSO As Object, ByVal Caminho As String)
    If Not FSO.FolderExists(Caminho) Then FSO.CreateFolder Caminho
End Sub

Private Sub CriarPastaRecursiva(ByVal FSO As Object, ByVal Caminho As String)
    Dim Parts As Variant, Partial As String, i As Long
    If FSO.FolderExists(Caminho) Then Exit Sub
    Parts = VBA.Split(Caminho, "\")
    Partial = Parts(0)
    For i = 1 To UBound(Parts)
        Partial = Partial & "\" & Parts(i)
        If Not FSO.FolderExists(Partial) Then
            On Error Resume Next
            FSO.CreateFolder Partial
            On Error GoTo 0
        End If
    Next i
End Sub
