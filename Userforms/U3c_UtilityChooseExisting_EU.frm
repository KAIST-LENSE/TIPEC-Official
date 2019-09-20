VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3c_UtilityChooseExisting_EU 
   Caption         =   "Choose from Existing Energy Utility Database"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   OleObjectBlob   =   "U3c_UtilityChooseExisting_EU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3c_UtilityChooseExisting_EU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3c_EU_Choose_Change()
' [1] Display Selected DB-Material Info
U3c_EU_SelectedTextBox.Text = U3c_EU_Choose.Column(0) & "   |   " & U3c_EU_Choose.Column(1)
End Sub


Private Sub U3c_Button_AddfromDB_EU_Click()
' [0] Check if less than 20 energy utilities already
If Worksheets("B3").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Energy Utilities already specified!! (20)")
    Exit Sub
End If

' [1] Declare Variables
Dim EUFoundDB As Range
Dim currentrowDB As Double

' [2] DB-Material Name as String. Find Cell with String Value and Get Row Index
EUNameDB = U3c_EU_Choose.Column(1)
Set EUFoundDB = Sheets("DB2").Range("C5:C2000").Find(What:=EUNameDB, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
currentrowDB = EUFoundDB.Row

' [3] Find last row in Current Project Material List
ActiveRow_UP = Sheets("B3").Range("B65000").End(xlUp).Row
ActiveRow = ActiveRow_UP + 1

' [4] Assign Row Index to a variable, current Row
Sheets("B3").Cells(ActiveRow, 2) = ActiveRow - 4
Sheets("B3").Cells(ActiveRow, 3) = Sheets("DB2").Cells(currentrowDB, 3)
Sheets("B3").Cells(ActiveRow, 4) = Sheets("DB2").Cells(currentrowDB, 4)
Sheets("B3").Cells(ActiveRow, 5) = Sheets("DB2").Cells(currentrowDB, 5)
' *** CALCULATING UPDATED COST BASED ON AVERAGED INFLATION FACTOR ***
Sheets("B3").Cells(ActiveRow, 6) = ((1 + 0.016) ^ ((Sheets("B1").Cells(5, 3)) - (Sheets("DB2").Cells(currentrowDB, 6)))) * Sheets("DB2").Cells(currentrowDB, 7)

' [5] Update Display Form with Added Entry
If Sheets("S2").Range("G17").Interior.Color = RGB(248, 203, 173) Then
    ' Mass Utility Index
    Set UtilityM_Index_UpTo20 = Sheets("B4").Range("B5:B24")
    Set DisplayM_Index_UpTo20 = Sheets("S2").Range("G15:G34")
    DisplayM_Index_UpTo20.Value = UtilityM_Index_UpTo20.Value
    
    ' Mass Utility Name
    Set UtilityM_Name_UpTo20 = Sheets("B4").Range("C5:C24")
    Set DisplayM_Name_UpTo20 = Sheets("S2").Range("H15:H34")
    DisplayM_Name_UpTo20.Value = UtilityM_Name_UpTo20.Value
        
    ' Mass Utility CO2 Footprint Production
    Set UtilityM_CO2P_UpTo20 = Sheets("B4").Range("D5:D24")
    Set DisplayM_CO2P_UpTo20 = Sheets("S2").Range("J15:J34")
    DisplayM_CO2P_UpTo20.Value = UtilityM_CO2P_UpTo20.Value
        
    ' Mass Utility CO2 Footprint Consumption
    Set UtilityM_CO2C_UpTo20 = Sheets("B4").Range("E5:E24")
    Set DisplayM_CO2C_UpTo20 = Sheets("S2").Range("K15:K34")
    DisplayM_CO2C_UpTo20.Value = UtilityM_CO2C_UpTo20.Value
        
    ' Mass Utility Specific Cost
    Set UtilityM_Cost_UpTo20 = Sheets("B4").Range("F5:F24")
    Set DisplayM_Cost_UpTo20 = Sheets("S2").Range("L15:L34")
    DisplayM_Cost_UpTo20.Value = UtilityM_Cost_UpTo20.Value
    
    Else
    ' Energy Utility Index
    Set UtilityE_Index_UpTo20 = Sheets("B3").Range("B5:B24")
    Set DisplayE_Index_UpTo20 = Sheets("S2").Range("G15:G34")
    DisplayE_Index_UpTo20.Value = UtilityE_Index_UpTo20.Value
    
    ' Energy Utility Name
    Set UtilityE_Name_UpTo20 = Sheets("B3").Range("C5:C24")
    Set DisplayE_Name_UpTo20 = Sheets("S2").Range("H15:H34")
    DisplayE_Name_UpTo20.Value = UtilityE_Name_UpTo20.Value
        
    ' Energy Utility CO2 Footprint Production
    Set UtilityE_CO2P_UpTo20 = Sheets("B3").Range("D5:D24")
    Set DisplayE_CO2P_UpTo20 = Sheets("S2").Range("J15:J34")
    DisplayE_CO2P_UpTo20.Value = UtilityE_CO2P_UpTo20.Value
        
    ' Energy Utility CO2 Footprint Consumption
    Set UtilityE_CO2C_UpTo20 = Sheets("B3").Range("E5:E24")
    Set DisplayE_CO2C_UpTo20 = Sheets("S2").Range("K15:K34")
    DisplayE_CO2C_UpTo20.Value = UtilityE_CO2C_UpTo20.Value
        
    ' Energy Utility Specific Cost
    Set UtilityE_Cost_UpTo20 = Sheets("B3").Range("F5:F24")
    Set DisplayE_Cost_UpTo20 = Sheets("S2").Range("L15:L34")
    DisplayE_Cost_UpTo20.Value = UtilityE_Cost_UpTo20.Value
End If

' [6] Display Material Successfully Updated
ans = MsgBox(EUNameDB & " has been added to project", vbOKOnly, "TIPEM- Utility Added")
If ans = vbYes Then
    U3c_UtilityChooseExisting_EU.Show
End If
End Sub


Private Sub U3c_ButtonEU_Close_Click()
' [1] Unload Userform
U3c_UtilityChooseExisting_EU.Hide
End Sub

Private Sub UserForm_Click()

End Sub
