VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U1a_DefineSpecies 
   Caption         =   "Define a New Species"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U1a_DefineSpecies.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U1a_DefineSpecies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***CLOSE USERFORM***
Private Sub U1a_Button_Cancel_Click()
Unload Me
End
End Sub



'***REGISTER A NEW SPECIES***
Private Sub U1a_Button_RegisterSpecies_Click()
If MsgBox("Register Species with the following Info??", vbYesNo) = vbYes Then
    ' [1] Assigns Entered Species Name to Database
    ThisWorkbook.Sheets("B1").Range("C8") = Me.U1a_Input_SpeciesName.Text
    ThisWorkbook.Sheets("S1").Range("N15") = Me.U1a_Input_SpeciesName.Text
    
    ' [2] Assigns Entered Species Origin to Database
    ThisWorkbook.Sheets("B1").Range("C9") = Me.U1a_ComboBox_SpeciesOrigin.Text
    ThisWorkbook.Sheets("S1").Range("N17") = Me.U1a_ComboBox_SpeciesOrigin.Text
    
    ' [3] Maximum Specific Growth Rate
    ThisWorkbook.Sheets("B1").Range("C10") = Me.U1a_Input_Growth.Text
    ThisWorkbook.Sheets("S1").Range("N22") = Me.U1a_Input_Growth.Text
    
    ' [4] Limiting Factor
    ThisWorkbook.Sheets("B1").Range("C11") = Me.U1a_Input_Limiting.Value
    ThisWorkbook.Sheets("S1").Range("N25") = Me.U1a_Input_Limiting.Value
    
    ' [5] Monod Half Saturation constant
    ThisWorkbook.Sheets("B1").Range("C12") = Me.U1a_Input_Monod.Value
    ThisWorkbook.Sheets("S1").Range("N28") = Me.U1a_Input_Monod.Value
Else
    Exit Sub
End If
End Sub
