VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3e_TransportAdd 
   Caption         =   "Add a New Transportation Method"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U3e_TransportAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3e_TransportAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3e_Button_ChooseExisting_Click()
' [1] Show Userform
U3g_TransportChooseExisting.Show
End Sub


Private Sub U3e_Button_TAdd_Click()
' [0] Check if less than 20 transportations already
TRANSPORT_Delete
Worksheets("S2").Activate
If Worksheets("B5").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Transportations already specified!! (20)")
    Exit Sub
End If

' [1] Add Registered Material to Last Row on Sheet
Dim LastRowData2 As Integer
LastRowData2 = Sheets("B5").Range("B65536").End(xlUp).Row
LastRowData2 = LastRowData2 + 1
Sheets("B5").Range("B" & LastRowData2) = LastRowData2 - 4
Sheets("B5").Range("C" & LastRowData2) = Me.U3e_TransportName.Text
Sheets("B5").Range("D" & LastRowData2) = Me.U3e_TCO2.Value
Sheets("B5").Range("E" & LastRowData2) = Me.U3e_TCost.Value

' [2] Re-Define Transportation Range
Dim Interval_T As String
Interval_T = "DB_Transportations_List"
ActiveWorkbook.Names.Add Name:=Interval_T, RefersToLocal:="=OFFSET('B5'!$B$4,1,,COUNTA('B5'!$C:$C),2)"

' [3] Redraw Transport Distance Matrix
TRANSPORT_Generate
Worksheets("S2").Activate

' [4] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [5] Clear Entries
Me.U3e_TransportName.Text = ""
Me.U3e_TCO2.Value = ""
Me.U3e_TCost.Value = ""
End Sub


Private Sub U3e_Button_TCancel_Click()
' [1] Display Transportations Data
Set Transport_Data = Sheets("B5").Range("B5:E24")
Set Transport_Display = Sheets("S2").Range("O15:R34")
Transport_Display.Value = Transport_Data.Value

' [2] Unload Userform
Unload Me
End
End Sub


Private Sub U3e_Label3_Click()

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

