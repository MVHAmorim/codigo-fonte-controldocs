Attribute VB_Name = "DocumentacaoControlDocs"
Option Explicit

#If VBA7 Then
    ' C�digo para sistemas de 32 e 64 bits no Office 2010 e posterior
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    ' C�digo para sistemas de 32 bits no Office 2007 e anteriores
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Function ControlPressionado() As Boolean
    ControlPressionado = (GetKeyState(vbKeyControl) And &H8000) <> 0
End Function

Public Function AcessarDocumentacao(ByRef control As IRibbonControl)

Dim URL As String
Dim urlBase As String

    urlBase = "https://controldocs-doc.escoladaautomacaofiscal.com.br/documentacao/"
    
    Select Case control.id
        
        'Bot�o Assinatura ControlDocs
        Case "btnAssinaturaControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs"
        
        '#Grupo Assinar ControlDocs
            
            '#Planos B�sicos
                'Bot�o Assinar Plano B�sico Mensal
                Case "btnBasicoMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-mensal-controldocs"
                
                'Bot�o Assinar Plano B�sico Semestral
                Case "btnBasicoSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-semestral-controldocs"
                
                'Bot�o Assinar Plano B�sico Anual
                Case "btnBasicoAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-basicos/plano-basico-anual-controldocs"
            
            '#Planos Plus
                'Bot�o Assinar Plano Plus Mensal
                Case "btnPlusMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-mensal-controldocs"
                
                'Bot�o Assinar Plano Plus Semestral
                Case "btnPlusSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-semestral-controldocs"
                
                'Bot�o Assinar Plano Plus Anual
                Case "btnPlusAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-plus/plano-plus-anual-controldocs"
            
            '#Planos Premium
                'Bot�o Assinar Plano Premium Mensal
                Case "btnPremiumMensal"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-mensal-controldocs"
                
                'Bot�o Assinar Plano Premium Semestral
                Case "btnPremiumSemestral"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-semestral-controldocs"
                
                'Bot�o Assinar Plano Premium Anual
                Case "btnPremiumAnual"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-assinar-controldocs/planos-Premium/plano-Premium-anual-controldocs"
            
            '#Assinatura Experimental
                'Bot�o Obter Assinatura Experimental
                Case "btnExperimentarControlDocs"
                    URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-assinar-controldocs/botao-obter-assinatura-experimental"
        
        '# Grupo Recursos de Assinatura
        
            'Bot�o Autenticar Usu�rio
            Case "btnAutenticarUsuario"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-autenticar-usuario"
        
            'Bot�o Consultar Assinatura
            Case "btnConsultarAssinatura"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-consultar-dados-da-assinatura"
        
            'Bot�o Limpar Dados Assinatura
            Case "btnLimparDadosAssinatura"
                URL = "guia-controldocs/grupo-navegacao-rapida/assinatura-controldocs/grupo-recursos-de-assinatura/botao-limpar-dados"
            
            
        'Bot�o Cadastro do Contribuinte
        Case "btnCadContrib"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte"
            
            'Bot�o Extrair Cadastro da Web
            Case "btnExtCadWeb"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte/extrair-cadastro-da-web"
                        
            'Bot�o Extrair Cadastro do SPED
            Case "btnExtCadSPED"
            URL = "guia-controldocs/grupo-navegacao-rapida/cadastro-do-contribuinte/extrair-cadastro-do-sped-fiscal"
                        
                        
        'Bot�o Recursos ControlDocs
        Case "btnRecursosControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs"
            
            'Bot�o Acessar Plataforma Educacional
            Case "btnControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/acessar-plataforma-educacional"
            
            'Bot�o Documenta��o ControlDocs
            Case "btnDocControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/documentacao-controldocs"
            
            'Bot�o Download da Vers�o Atual
            Case "btnDonwloadControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/download-da-versao-atual"
            
            'Bot�o Suporte Via WhatsApp
            Case "btnSuporte"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/suporte-via-whatsapp"
            
            'Bot�o Sugerir Melhorias
            Case "btnSugestoes"
            URL = "guia-controldocs/grupo-navegacao-rapida/recursos-controldocs/grupo-recursos-de-ajuda/sugerir-melhorias"
            
            
        'Bot�o Configura��es e Personaliza��es
        Case "btnConfiguracoesControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes"
            
            'CheckBox Remover Linhas de Grade
            Case "chLinhasGrade"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-personalizacao/remover-linhas-de-grade"
            
            'Bot�o Resetar ControlDocs
            Case "btnImportarExcel"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-configuracoes/botao-importar-dados-versao-anterior"
            
            'Bot�o Resetar ControlDocs
            Case "btnResetarControlDocs"
            URL = "guia-controldocs/grupo-navegacao-rapida/configuracoes-e-personalizacoes/grupo-configuracoes/resetar-controldocs"
            
            
        Case Else
            Call Util.MsgAviso("Esse recurso ainda n�o foi documentado." & vbCrLf & _
                "Caso precise de informa��es contate nosso suporte.", "Documenta��o ControlDocs")
            Exit Function
            
    End Select
    
    Call FuncoesLinks.AbrirUrl(urlBase & URL)
    
End Function
