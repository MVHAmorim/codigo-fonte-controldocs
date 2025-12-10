VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "relGestaoAssinatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Target.AddressLocal = "$B$2" Then
        
        With Target
            .Hyperlinks.Delete
            .Font.name = "Calibri"
            .Interior.Color = RGB(255, 255, 204)
        End With
        
    End If
    
End Sub
