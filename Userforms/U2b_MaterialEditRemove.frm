VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U2b_MaterialEditRemove 
   Caption         =   "Edit or Remove a Material"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U2b_MaterialEditRemove.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U2b_MaterialEditRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U2b_MaterialList_Change()
' [1] Display Selected Material Info
U2b_DisplayMaterial.Text = U2b_MaterialList.Column(0) & "   |   " & U2b_MaterialList.Column(1)

' [2] Material Selected from List
If Me.U2b_MaterialList.Value = "" Then
MsgBox "You must specify a Material to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Material Selected as Indexing String
MatName = U2b_MaterialList.Column(0)
On Error Resume Next

' [4] Find the Row of the Indexed Material, then Load Values
Me.U2b_Show_Country.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 3, 0)
Me.U2b_Show_Year.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 4, 0)
Me.U2b_Show_CO2Prod.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 5, 0)
Me.U2b_Show_CO2Cons.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 6, 0)
Me.U2b_Show_Purchase.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 7, 0)
Me.U2b_Show_Selling.Value = Application.WorksheetFunction.VLookup(MatName, Sheets("B2").Range("B4:I2000"), 8, 0)
End Sub


Private Sub U2b_Button_RemoveMaterial_Click()
' [1] Declare Variables
On Error Resume Next
Dim msg As String
Dim ans As String
Dim rowselect As Double
Dim MatFound As Range

' [2] Must Select Material Before Proceeding
If Me.U2b_MaterialList.Value = "" Then
MsgBox "You must specify a Material to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Material Name as Double. Find Cell with Value and Get Row Index
MatName2 = U2b_MaterialList.Column(1)
Set MatFound = Sheets("B2").Range("C4:C2000").Find(What:=MatName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Move Row Index After Material has been Deleted
If MsgBox("Are you sure you want to delete " & MatName2 & "?", vbYesNo) = vbNo Then Exit Sub
rowselect = MatFound.Row
Sheets("B2").Rows(rowselect).EntireRow.Delete
rowselect = rowselect - 1

' [5] Re-Define Materials Range
Dim Interval_M As String
Interval_M = "DB_MaterialsList"
ActiveWorkbook.Names.Add Name:=Interval_M, RefersToLocal:="=OFFSET('B2'!$B$4,0,,COUNTA('B2'!$C:$C),2)"

' [6] Auto-Re-Number the Material Name Index After Deleting Row
Dim Count As Integer
Dim NumbMat As Integer
Count = 1
NumbMat = Sheets("B2").Range("K3").Value
Application.EnableEvents = False
Do While Count <= NumbMat
    Sheets("B2").Range("B" & Count + 3).Value = Count
    Count = Count + 1
Loop
Sheets("B2").Range("A" & NumbMat + 1).Value = ""
Application.EnableEvents = True

' [7] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
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

' [8] Show Msg that Material Was Deleted
ans = MsgBox(MatName2 & " has been successfully deleted", vbOKOnly, "TIPEM- Material Deleted")
If ans = vbYes Then
    U2b_EditRemoveMaterial.Show
End If
End Sub


Private Sub U2b_Button_UpdateMaterial_Click()
On Error Resume Next
' [1] Must select a material to proceed updating
If Me.U2b_MaterialList.Value = "" Then
MsgBox "You must specify a Material to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [2] Declare Variables
Dim MatFound As Range
Dim currentrow As Double

' [3] Material Name as String. Find Cell with String Value and Get Row Index
MatName3 = U2b_MaterialList.Column(0)
Set MatFound = Sheets("B2").Range("B4:B2000").Find(What:=MatName3, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Assign Row Index to a variable, current Row
currentrow = MatFound.Row
Sheets("B2").Cells(currentrow, 4) = Me.U2b_Show_Country.Value
Sheets("B2").Cells(currentrow, 5) = Me.U2b_Show_Year.Value
Sheets("B2").Cells(currentrow, 6) = Me.U2b_Show_CO2Prod.Value
Sheets("B2").Cells(currentrow, 7) = Me.U2b_Show_CO2Cons.Value
Sheets("B2").Cells(currentrow, 8) = Me.U2b_Show_Purchase.Value
Sheets("B2").Cells(currentrow, 9) = Me.U2b_Show_Selling.Value

' [5] Re-Define Materials Range
Dim Interval_M As String
Interval_M = "DB_MaterialsList"
ActiveWorkbook.Names.Add Name:=Interval_M, RefersToLocal:="=OFFSET('B2'!$B$4,0,,COUNTA('B2'!$C:$C),2)"

' [6] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
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

' [7] Display Material Successfully Updated
ans = MsgBox(MatName3 & " has been successfully updated", vbOKOnly, "TIPEM- Material Updated")
If ans = vbYes Then
    U2b_EditRemoveMaterial.Show
End If
End Sub


Private Sub U2b_Button_Cancel_Click()
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


Private Sub UserForm_Initialize()
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
