VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U7d_DCFROR 
   Caption         =   "Discounted Cash Flow Analysis"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5445
   OleObjectBlob   =   "U7d_DCFROR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U7d_DCFROR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** PERFORM DCFROR ****
Private Sub U7d_DCFROR_Click()
' [1] Warning: Perform DCFROR?
    If MsgBox("Perform DCFROR Analysis??", vbYesNo) = vb Then
        Exit Sub
    Else
        MsgBox "DCFROR Results are Available", vbExclamation, "TIPEM- Notice"
    End If
End Sub

Private Sub UserForm_Click()

End Sub
