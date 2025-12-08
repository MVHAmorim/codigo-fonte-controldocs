VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportadorRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SeletorRegistros As ExportadorRegistros_ListaBlocos
Private colPlanilhas As Collection
Private colRegistros As Collection

Public Function ExportarRegistros(ParamArray Registros() As Variant)

Dim i As Long
Dim Plan As Worksheet
Dim Registro As Dictionary
    
    Call InicializarObjetos
    
    Call SeletorRegistros.CarregarRegistrosExportacao(Registros)
    Set colPlanilhas = SeletorRegistros.colPlanilhas
    Set colRegistros = SeletorRegistros.colRegistros
    
    For i = 1 To colPlanilhas.Count
        
        Set Plan = colPlanilhas.Item(i)
        Set Registro = colRegistros.Item(i)
        
        If Not Registro Is Nothing Then
            
            Call Util.AtualizarBarraStatus("Exportando dados do registro " & Plan.name)
            
            Call Util.LimparDados(Plan, 4, False)
            Call Util.ExportarDadosDicionario(Plan, Registro, "A4")
            
        End If
        
    Next i
    
    Set Plan = Nothing
    Set Registro = Nothing
    
    Call LimparObjetos
    
End Function

Private Sub InicializarObjetos()
    
    Set SeletorRegistros = New ExportadorRegistros_ListaBlocos
    
End Sub

Private Sub LimparObjetos()
    
    Set SeletorRegistros = New ExportadorRegistros_ListaBlocos
    Set colPlanilhas = New Collection
    Set colRegistros = New Collection
    
    Call Util.AtualizarBarraStatus(False)
    
End Sub

