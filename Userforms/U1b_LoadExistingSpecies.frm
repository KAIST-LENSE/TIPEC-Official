VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U1b_LoadExistingSpecies 
   Caption         =   "Choose an Existing Species"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10485
   OleObjectBlob   =   "U1b_LoadExistingSpecies.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U1b_LoadExistingSpecies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***CANCEL***
Private Sub U1b_Cancel_Click()
Unload Me
End
End Sub




'***ENTRY SELECTED***
Private Sub U1b_Combobox_Change()
' [1] Species Selected from Database
    If Me.U1b_Combobox.Value = "" Then
    MsgBox "Please select an entry from the Database!!!", vbExclamation, "TIPEM- Error"
    Exit Sub
    End If

' [2] Display Selected Material Info
    U1b_Display.Text = U1b_Combobox.Column(1) & "   |      " & U1b_Combobox.Column(2) & " day^-1          " & U1b_Combobox.Column(3) & " Limiting          " & U1b_Combobox.Column(4) & "kg/m^3"
End Sub




'***REGISTER SPECIES***
Private Sub U1b_Ok_Click()
' [1] Register to Display Sheet and Backend Sheet
    ' [1] Assigns Entered Species Name to Database
    ThisWorkbook.Sheets("B1").Range("C8") = Me.U1b_Combobox.Column(1)
    ThisWorkbook.Sheets("S1").Range("N15") = Me.U1b_Combobox.Column(1)
    
    ' [2] Assigns Entered Species Origin to Database
    ThisWorkbook.Sheets("B1").Range("C9") = Worksheets("B1").Range("C4").Value
    ThisWorkbook.Sheets("S1").Range("N17") = Worksheets("B1").Range("C4").Value
    
    ' [3] Maximum Specific Growth Rate
    ThisWorkbook.Sheets("B1").Range("C10") = Me.U1b_Combobox.Column(2)
    ThisWorkbook.Sheets("S1").Range("N22") = Me.U1b_Combobox.Column(2)
    
    ' [4] Limiting Factor
    ThisWorkbook.Sheets("B1").Range("C11") = Me.U1b_Combobox.Column(3)
    ThisWorkbook.Sheets("S1").Range("N25") = Me.U1b_Combobox.Column(3)
    
    ' [5] Monod Half Saturation constant
    ThisWorkbook.Sheets("B1").Range("C12") = Me.U1b_Combobox.Column(4)
    ThisWorkbook.Sheets("S1").Range("N28") = Me.U1b_Combobox.Column(4)
End Sub

Private Sub UserForm_Click()

End Sub
