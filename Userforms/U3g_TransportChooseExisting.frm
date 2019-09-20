VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3g_TransportChooseExisting 
   Caption         =   "Choose from Existing Transportation Database"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   OleObjectBlob   =   "U3g_TransportChooseExisting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3g_TransportChooseExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3f_TransportationChoose_Change()
' [1] Display Selected DB-Material Info
U3f_Selected_Transport.Text = U3f_TransportationChoose.Column(0) & "   |   " & U3f_TransportationChoose.Column(1)
End Sub


Private Sub U3f_Button_AddfromDB_Click()
' [0] Check if less than 20 transportations already
TRANSPORT_Delete
Worksheets("S2").Activate
If Worksheets("B5").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Transportations already specified!! (20)")
    Exit Sub
End If

' [1] Declare Variables
Dim TFoundDB As Range
Dim currentrowDB As Double

' [2] DB-Material Name as String. Find Cell with String Value and Get Row Index
TNameDB = U3f_TransportationChoose.Column(1)
Set TFoundDB = Sheets("DB3").Range("C4:C2000").Find(What:=TNameDB, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
currentrowDB = TFoundDB.Row

' [3] Find last row in Current Project Material List
ActiveRow_UP = Sheets("B5").Range("B65000").End(xlUp).Row
ActiveRow = ActiveRow_UP + 1

' [4] Assign Row Index to a variable, current Row
Sheets("B5").Cells(ActiveRow, 2) = ActiveRow - 4
Sheets("B5").Cells(ActiveRow, 3) = Sheets("DB3").Cells(currentrowDB, 3)
Sheets("B5").Cells(ActiveRow, 4) = Sheets("DB3").Cells(currentrowDB, 4)
' *** CALCULATING UPDATED COST BASED ON AVERAGED INFLATION FACTOR ***
Sheets("B5").Cells(ActiveRow, 5) = ((1 + 0.016) ^ ((Sheets("B1").Cells(5, 3)) - (Sheets("DB3").Cells(currentrowDB, 5)))) * Sheets("DB3").Cells(currentrowDB, 6)

' [5] Re-Define Transportation Range
Dim Interval_T As String
Interval_T = "DB_Transportations_List"
ActiveWorkbook.Names.Add Name:=Interval_T, RefersToLocal:="=OFFSET('B5'!$B$4,1,,COUNTA('B5'!$C:$C),2)"

' [6] Redraw Transportation Distance Matrix
TRANSPORT_Generate
Worksheets("S2").Activate

' [7] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [8] Display Material Successfully Updated
ans = MsgBox(TNameDB & " has been added to project", vbOKOnly, "TIPEM- Transportation Added")
If ans = vbYes Then
    U3g_TransportChooseExisting.Show
End If
End Sub


Private Sub U3f_Button_Close_Click()
' [5] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [2] Unload Userform
U3g_TransportChooseExisting.Hide
End Sub

