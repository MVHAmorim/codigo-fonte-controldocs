VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomPartXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SalvarJson(ByVal Namespace As String, ByVal jsonString As String)

Dim XML As String, TagRaiz$
    
    TagRaiz = VBA.Split(Namespace, "_")(0)
    
    XML = "<" & TagRaiz & " xmlns='" & Namespace & "'>"
    XML = XML & "<json><![CDATA[" & jsonString & "]]></json>"
    XML = XML & "</" & TagRaiz & ">"
    
    Call DeletarXmlPart(Namespace)
    Call ThisWorkbook.CustomXMLParts.Add(XML)
    
End Sub

Public Sub SalvarTabelaTXT(ByVal Namespace As String, ByVal DadosTabela As String)

Dim XML As String, TagRaiz$
    
    TagRaiz = "TabelaTXT"
    
    XML = "<" & TagRaiz & " xmlns='" & Namespace & "'>"
    XML = XML & "<txt><![CDATA[" & DadosTabela & "]]></txt>"
    XML = XML & "</" & TagRaiz & ">"
    
    Call DeletarXmlPart(Namespace)
    Call ThisWorkbook.CustomXMLParts.Add(XML)
    
End Sub

Public Function ExtrairJsonXmlPart(ByVal Namespace As String) As String

Dim xmlPartsExistentes As Office.CustomXMLParts
Dim XmlPart As Office.CustomXMLPart
Dim Part As Office.CustomXMLPart
Dim nodeJson As IXMLDOMNode
Dim XML As DOMDocument60
Dim TagRaiz As String
    
    Set xmlPartsExistentes = ThisWorkbook.CustomXMLParts.SelectByNamespace(Namespace)
    For Each XmlPart In xmlPartsExistentes
        
        Set XML = fnXML.RemoverNamespaces(XmlPart.XML, True)
        
        TagRaiz = VBA.Split(Namespace, "_")(0)
        If VBA.LCase(XML.DocumentElement.BaseName) = VBA.LCase(TagRaiz) Then
            
            Set nodeJson = XML.DocumentElement.SelectSingleNode("json")
            If Not nodeJson Is Nothing Then ExtrairJsonXmlPart = nodeJson.text
            Exit Function
            
        End If
        
    Next XmlPart
    
End Function

Public Function ExtrairTXTPartXML(ByVal Namespace As String) As String

Dim xmlPartsExistentes As Office.CustomXMLParts
Dim XmlPart As Office.CustomXMLPart
Dim Part As Office.CustomXMLPart
Dim nodeTXT As IXMLDOMNode
Dim XML As DOMDocument60
Dim TagRaiz As String
    
    Set xmlPartsExistentes = ThisWorkbook.CustomXMLParts.SelectByNamespace(Namespace)
    For Each XmlPart In xmlPartsExistentes
        
        Set XML = fnXML.RemoverNamespaces(XmlPart.XML, True)
        
        TagRaiz = "TabelaTXT"
        If VBA.LCase(XML.DocumentElement.BaseName) = VBA.LCase(TagRaiz) Then
            
            Set nodeTXT = XML.DocumentElement.SelectSingleNode("txt")
            If Not nodeTXT Is Nothing Then ExtrairTXTPartXML = nodeTXT.text
            Exit Function
            
        End If
        
    Next XmlPart
    
End Function

Private Sub DeletarXmlPart(ByVal Namespace As String)

Dim xmlPartsExistentes As Office.CustomXMLParts
Dim XmlPart As Office.CustomXMLPart
    
    Set xmlPartsExistentes = ThisWorkbook.CustomXMLParts.SelectByNamespace(Namespace)
    For Each XmlPart In xmlPartsExistentes
        
        XmlPart.Delete
        
    Next XmlPart
    
End Sub
