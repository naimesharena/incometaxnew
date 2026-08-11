Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{D7969338-4365-4BDC-9A69-3C358F3C36EC}{371E17A1-9648-48E0-93C3-22461A0C0F76}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
    Cancel = True
End If
End Sub
