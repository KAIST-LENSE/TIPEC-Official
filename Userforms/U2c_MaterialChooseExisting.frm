VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U2c_MaterialChooseExisting 
   Caption         =   "Choose from Existing Material Database"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   OleObjectBlob   =   "U2c_MaterialChooseExisting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U2c_MaterialChooseExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U2c_MaterialsChoose_Change()
' [1] Display Selected DB-Material Info
U2c_SelectedTextBox.Text = U2c_MaterialsChoose.Column(0) & "   |   " & U2c_MaterialsChoose.Column(1)
End Sub


Private Sub U2c_Button_AddfromDB_Click()
' [1] Declare Variables
Dim MatFoundDB As Range
Dim currentrowDB As Double

' [2] DB-Material Name as String. Find Cell with String Value and Get Row Index
MatNameDB = U2c_MaterialsChoose.Column(1)
Set MatFoundDB = Sheets("DB1").Range("C4:C2000").Find(What:=MatNameDB, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
currentrowDB = MatFoundDB.Row

' [3] Find last row in Current Project Material List
ActiveRow_UP = Sheets("B2").Range("B65000").End(xlUp).Row
ActiveRow = ActiveRow_UP + 1

' [4] Assign Row Index to a variable, current Row
Sheets("B2").Cells(ActiveRow, 2) = ActiveRow - 3
Sheets("B2").Cells(ActiveRow, 3) = Sheets("DB1").Cells(currentrowDB, 3)
Sheets("B2").Cells(ActiveRow, 4) = Sheets("DB1").Cells(currentrowDB, 4)
Sheets("B2").Cells(ActiveRow, 5) = Sheets("DB1").Cells(currentrowDB, 5)
Sheets("B2").Cells(ActiveRow, 6) = Sheets("DB1").Cells(currentrowDB, 6)
Sheets("B2").Cells(ActiveRow, 7) = Sheets("DB1").Cells(currentrowDB, 7)
Sheets("B2").Cells(ActiveRow, 8) = Sheets("DB1").Cells(currentrowDB, 8)
Sheets("B2").Cells(ActiveRow, 9) = Sheets("DB1").Cells(currentrowDB, 9)

' [5] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
If Sheets("B2").Range("K3") >= 21 Then
    S1.ScrollBar2.Visible = True
    S1.ScrollBar2.Max = S3_2.UsedRange.Rows.Count - 19
    S1.ScrollBar2.Min = 4
    S1.ScrollBar2.Value = 5
    Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
    Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
    DisplayUpTo20.Value = MaterialsUpTo20.Value
    Else
        S1.ScrollBar2.Visible = False
        Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
        Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
        DisplayUpTo20.Value = MaterialsUpTo20.Value
End If

' [6] Display Material Successfully Updated
ans = MsgBox(MatNameDB & " has been added to project", vbOKOnly, "TIPEM- Material Added")
If ans = vbYes Then
    U2c_MaterialChooseExisting.Show
End If
End Sub


Private Sub U2c_Button_Close_Click()
' [1] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
If Sheets("B2").Range("K3") >= 21 Then
    S1.ScrollBar2.Visible = True
    S1.ScrollBar2.Max = S3_2.UsedRange.Rows.Count - 19
    S1.ScrollBar2.Min = 4
    S1.ScrollBar2.Value = 5
    Else
        S1.ScrollBar2.Visible = False
        Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
        Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
        DisplayUpTo20.Value = MaterialsUpTo20.Value
End If

' [2] Unload Userform
U2c_MaterialChooseExisting.Hide
End Sub

Private Sub UserForm_Click()

End Sub
