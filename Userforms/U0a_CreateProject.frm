VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U0a_CreateProject 
   Caption         =   "Create a New TIPEM Project"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U0a_CreateProject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U0a_CreateProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U0a_Button_Ok_Click()
' [1] Assign Project Information to Datasheet
ThisWorkbook.Sheets("B1").Range("C3") = Me.U0a_Input_ProjectName.Text
ThisWorkbook.Sheets("B1").Range("C4") = Me.U0a_ComboBox_PlantLocation.Text
ThisWorkbook.Sheets("B1").Range("C5") = Me.U0a_Input_EvaluatorName.Text

' [2]Unload Userform
Worksheets("S1").Activate
ActiveWindow.Zoom = 110
Range("A1").Select
Unload Me
End Sub


Private Sub U0a_Button_Cancel_Click()
' [1] Unload Userform
Unload Me
End
End Sub

Private Sub UserForm_Click()

End Sub
