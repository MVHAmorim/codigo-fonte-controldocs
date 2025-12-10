VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutorMetodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ExecutarMetodo(ByVal NomeMetodo As String, ByVal Processador As Iprocessador, ByVal TipoSPED As String, ByVal Registro As String)
    
    On Error GoTo TratarErro:
    
    Call Processador.Executar(TipoSPED, Registro)
    Set Processador = Nothing
    
Exit Sub
TratarErro:

Dim infoErro As New clsGerenciadorErros
    
    Call infoErro.NotificarErroInesperado(TypeName(Processador), NomeMetodo)
    
End Sub
