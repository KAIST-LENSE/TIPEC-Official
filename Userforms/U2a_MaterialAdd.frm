VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U2a_MaterialAdd 
   Caption         =   "Add a New Material"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U2a_MaterialAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U2a_MaterialAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U2a_DefineMaterial_Click()
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
End Sub


Private Sub U2a_Button_ChooseExisting_Click()
' [1] Show Userform
U2c_MaterialChooseExisting.Show
End Sub


Private Sub U2a_Button_MaterialRegister_Click()
' [0] Check that the entered material has no spaces
If InStr(U2a_Input_MaterialName, " ") > 0 Then
    MsgBox "Material Name must be unique with no spaces or special characters!!!", vbExclamation, "TIPEM- Error"
    Exit Sub
End If

' [1] Add Registered Material to Last Row on Sheet
Dim LastRowData As Integer
LastRowData = Sheets("B2").Range("B65536").End(xlUp).Row
LastRowData = LastRowData + 1

' [2] Update Materials Database Sheet (B2) with the Materials Info from Userform
Sheets("B2").Range("B" & LastRowData) = LastRowData - 3
Sheets("B2").Range("C" & LastRowData) = Me.U2a_Input_MaterialName.Text
Sheets("B2").Range("D" & LastRowData) = Me.U2a_Check_Country.Value
Sheets("B2").Range("E" & LastRowData) = Me.U2a_Check_Year.Value
Sheets("B2").Range("F" & LastRowData) = Me.U2a_Input_CO2Prod.Value
Sheets("B2").Range("G" & LastRowData) = Me.U2a_Input_CO2Cons.Value
Sheets("B2").Range("H" & LastRowData) = Me.U2a_Input_Purchase.Value
Sheets("B2").Range("I" & LastRowData) = Me.U2a_Input_Selling.Value

' [3] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
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

' [4] Clear Entries
Me.U2a_Input_MaterialName.Text = ""
Me.U2a_Check_Country.Value = ""
Me.U2a_Check_Year.Value = ""
Me.U2a_Input_CO2Prod.Value = ""
Me.U2a_Input_CO2Cons.Value = ""
Me.U2a_Input_Purchase.Value = ""
Me.U2a_Input_Selling.Value = ""

' [5] Re-Define Materials Range
Dim Interval_M As String
Interval_M = "DB_MaterialsList"
ActiveWorkbook.Names.Add Name:=Interval_M, RefersToLocal:="=OFFSET('B2'!$B$4,0,,COUNTA('B2'!$C:$C),2)"
End Sub


Private Sub U2a_Button_Cancel_Click()
On Error Resume Next
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

' [2] Re-Define Materials Range
Dim Interval_M As String
Interval_M = "DB_MaterialsList"
ActiveWorkbook.Names.Add Name:=Interval_M, RefersToLocal:="=OFFSET('B2'!$B$4,0,,COUNTA('B2'!$C:$C),2)"

' [3] Unload Userform
Unload Me
End
End Sub


Private Sub UserForm_Click()
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
End Sub
