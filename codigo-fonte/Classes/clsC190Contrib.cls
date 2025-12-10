VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsC190Contrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Processador As Iprocessador
Private Executor As New clsExecutorMetodos

Public Sub AtualizarNCM_C190()
    
    Inicio = Now()
    Call Executor.ExecutarMetodo("AtualizarNCM_C190", New clsC190Contrib_AtualizarNCM, "Contribuições", "C190")
    
End Sub

Public Sub AgruparRegistros()
    
    Inicio = Now()
    Call Executor.ExecutarMetodo("AgruparRegistros", New fnSPED_AgruparRegistros, "Contribuições", "C190")
    
End Sub
