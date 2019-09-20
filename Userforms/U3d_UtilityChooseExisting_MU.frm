VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3d_UtilityChooseExisting_MU 
   Caption         =   "Choose from Existing Mass Utility Database"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   OleObjectBlob   =   "U3d_UtilityChooseExisting_MU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3d_UtilityChooseExisting_MU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3d_MU_Choose_Change()
' [1] Display Selected DB-Material Info
U3d_MU_SelectedTextBox.Text = U3d_MU_Choose.Column(0) & "   |   " & U3d_MU_Choose.Column(1)
End Sub


Private Sub U3d_Button_AddfromDB_MU_Click()
' [0] Check if less than 20 mass utilities already
If Worksheets("B4").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Mass Utilities already specified!! (20)")
    Exit Sub
End If

' [1] Declare Variables
Dim MUFoundDB As Range
Dim currentrowDB As Double

' [2] DB-Material Name as String. Find Cell with String Value and Get Row Index
MUNameDB = U3d_MU_Choose.Column(1)
Set MUFoundDB = Sheets("DB2").Range("K5:K2000").Find(What:=MUNameDB, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
currentrowDB = MUFoundDB.Row

' [3] Find last row in Current Project Material List
ActiveRow_UP = Sheets("B4").Range("B65000").End(xlUp).Row
ActiveRow = ActiveRow_UP + 1

' [4] Assign Row Index to a variable, current Row
Sheets("B4").Cells(ActiveRow, 2) = ActiveRow - 4
Sheets("B4").Cells(ActiveRow, 3) = Sheets("DB2").Cells(currentrowDB, 11)
Sheets("B4").Cells(ActiveRow, 4) = Sheets("DB2").Cells(currentrowDB, 12)
Sheets("B4").Cells(ActiveRow, 5) = Sheets("DB2").Cells(currentrowDB, 13)
' *** CALCULATING UPDATED COST BASED ON AVERAGED INFLATION FACTOR ***
Sheets("B4").Cells(ActiveRow, 6) = ((1 + 0.016) ^ ((Sheets("B1").Cells(5, 3)) - (Sheets("DB2").Cells(currentrowDB, 14)))) * Sheets("DB2").Cells(currentrowDB, 15)

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
ans = MsgBox(MUNameDB & " has been added to project", vbOKOnly, "TIPEM- Utility Added")
If ans = vbYes Then
    U3d_UtilityChooseExisting_MU.Show
End If
End Sub


Private Sub U3d_ButtonMU_Close_Click()
' [1] Unload Userform
U3d_UtilityChooseExisting_MU.Hide
End Sub
