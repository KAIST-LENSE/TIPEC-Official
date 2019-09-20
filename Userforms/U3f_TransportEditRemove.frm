VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3f_TransportEditRemove 
   Caption         =   "Select a Transportation to Edit"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U3f_TransportEditRemove.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3f_TransportEditRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3f_TransportList_Change()
' [1] Display Selected Material Info
U3f_Display_Transport.Text = U3f_TransportList.Column(0) & "   |   " & U3f_TransportList.Column(1)

' [1] Transportation Selected from List
If Me.U3f_TransportList.Value = "" Then
MsgBox "You must specify a Transportation to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [2] Transportation Selected as Indexing String
TName = U3f_TransportList.Column(0)
On Error Resume Next

' [3] Find the Row of the Indexed Transportation, then Load Values
Me.U3f_Show_TCO2.Value = Application.WorksheetFunction.VLookup(TName, Sheets("B5").Range("B5:E2000"), 3, 0)
Me.U3f_Show_TCost.Value = Application.WorksheetFunction.VLookup(TName, Sheets("B5").Range("B5:E2000"), 4, 0)
End Sub


Private Sub U3f_Button_RemoveT_Click()
On Error Resume Next
' [0] Delete Transportation Distance Matrix
TRANSPORT_Delete
Worksheets("S2").Activate

' [1] Declare Variables
Dim msg As String
Dim ans As String
Dim rowselect As Double
Dim TFound As Range

' [2] Must Select Transportation Before Proceeding
If Me.U3f_TransportList.Value = "" Then
MsgBox "You must specify a Transportation to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Transportation Name as String. Find Cell with String Value and Get Row Index
TName2 = U3f_TransportList.Column(1)
On Error Resume Next
Set TFound = Sheets("B5").Range("B5:E2000").Find(What:=TName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Move Row Index After Transportation has been Deleted
If MsgBox("Are you sure you want to delete " & TName2 & "?", vbYesNo) = vbNo Then Exit Sub
rowselect = TFound.Row
Sheets("B5").Rows(rowselect).EntireRow.Delete
rowselect = rowselect - 1

' [5] Re-Define Transportation Range
Dim Interval_T As String
Interval_T = "DB_Transportations_List"
ActiveWorkbook.Names.Add Name:=Interval_T, RefersToLocal:="=OFFSET('B5'!$B$4,1,,COUNTA('B5'!$C:$C),2)"

' [6] Auto-Re-Number the Material Name Index After Deleting Row
Dim Count As Integer
Dim NumbMat As Integer
Count = 1
NumbMat = Sheets("B5").Range("C1").Value
Application.EnableEvents = False
Do While Count <= NumbMat
    Sheets("B5").Range("B" & Count + 4).Value = Count
    Count = Count + 1
Loop
Sheets("B5").Range("A" & NumbMat + 1).Value = ""
Application.EnableEvents = True

' [7] Generate Transport Distance Matrix
TRANSPORT_Generate
Worksheets("S2").Activate

' [8] Display Transportation Data on Display
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [9] Show Msg that Transportation Was Deleted
ans = MsgBox(TName2 & " has been successfully deleted", vbOKOnly, "TIPEM- Transportation Mode Deleted")
If ans = vbYes Then
    U3f_TransportEditRemove.Show
End If
End Sub


Private Sub U3f_Button_UpdateT_Click()
On Error Resume Next
' [1] Must select a Transportation to proceed updating
If Me.U3f_TransportList.Value = "" Then
MsgBox "You must specify a Transportation to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [2] Declare Variables
Dim TFound As Range
Dim currentrow As Double

' [3] Transportation Name as String. Find Cell with String Value and Get Row Index
TName2 = U3f_TransportList.Column(0)
Set TFound = Sheets("B5").Range("B5:B2000").Find(What:=TName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Assign Row Index to a variable, current Row
currentrow = TFound.Row
Sheets("B5").Cells(currentrow, 4) = U3f_Show_TCO2.Value
Sheets("B5").Cells(currentrow, 5) = U3f_Show_TCost.Value
    
' [5] Display Transportation Data to Display
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [6] Display Transportation Successfully Updated
ans = MsgBox(TName2 & " has been successfully updated", vbOKOnly, "TIPEM- Transportation Mode Updated")
If ans = vbYes Then
    U3f_TransportEditRemove.Show
End If
End Sub


Private Sub U3f_Button_Cancel_Click()
' [1] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [2] Unload Userform
Unload Me
End
End Sub


Private Sub UserForm_Click()
' [1] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value
End Sub


Private Sub UserForm_Initialize()
' [1] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value
End Sub
