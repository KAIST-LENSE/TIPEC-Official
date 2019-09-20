VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3a_UtilityAdd 
   Caption         =   "Add a New Utility"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U3a_UtilityAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3a_UtilityAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3a_Button_EUAdd_Click()
' [0] Check if less than 20 energy utilities already
If Worksheets("B3").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Energy Utilities already specified!! (20)")
    Exit Sub
End If

' [1] Add Registered Energy Utility to Last Row on Sheet
Dim EU_Sheet2 As Worksheet
Dim LastRowData2 As Integer
Set EU_Sheet2 = Sheets("B3")

LastRowData2 = EU_Sheet2.Range("B65536").End(xlUp).Row
LastRowData2 = LastRowData2 + 1
Sheets("B3").Range("B" & LastRowData2) = LastRowData2 - 4
Sheets("B3").Range("C" & LastRowData2) = Me.U3a_Input_EUName.Text
Sheets("B3").Range("D" & LastRowData2) = Me.U3a_Input_EUCO2Prod.Value
Sheets("B3").Range("E" & LastRowData2) = Me.U3a_Input_EUCO2Cons.Value
Sheets("B3").Range("F" & LastRowData2) = Me.U3a_Input_EUCost.Value

' [2] Re-Define Energy Utilities Range
Dim Interval_EU As String
Interval_EU = "DB_EUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_EU, RefersToLocal:="=OFFSET('B3'!$B$4,1,,COUNTA('B3'!$C:$C),2)"

' [3] Update Display Form with Added Entry
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

' [4] Clear Userform Entries for Next
Me.U3a_Input_EUName.Text = ""
Me.U3a_Input_EUCO2Prod.Value = ""
Me.U3a_Input_EUCO2Cons.Value = ""
Me.U3a_Input_EUCost.Value = ""
End Sub


Private Sub U3a_Button_MUAdd_Click()
' [0] Check if less than 20 mass utilities already
If Worksheets("B4").Range("C1").Value = 20 Then
    MsgBox ("Maximum number of Mass Utilities already specified!! (20)")
    Exit Sub
End If

' [1] Add Registered Mass Utility to Last Row on Sheet
Dim MU_Sheet3 As Worksheet
Dim LastRowData3 As Integer
Set MU_Sheet3 = Sheets("B4")

LastRowData3 = MU_Sheet3.Range("B65536").End(xlUp).Row
LastRowData3 = LastRowData3 + 1
Sheets("B4").Range("B" & LastRowData3) = LastRowData3 - 4
Sheets("B4").Range("C" & LastRowData3) = Me.U3a_Input_MUName.Text
Sheets("B4").Range("D" & LastRowData3) = Me.U3a_Input_MUCO2Prod.Value
Sheets("B4").Range("E" & LastRowData3) = Me.U3a_Input_MUCO2Cons.Value
Sheets("B4").Range("F" & LastRowData3) = Me.U3a_Input_MUCost.Value

' [2] Re-Define Mass Utilities Range
Dim Interval_MU As String
Interval_MU = "DB_MUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_MU, RefersToLocal:="=OFFSET('B4'!$B$4,1,,COUNTA('B4'!$C:$C),2)"

' [3] Update Display Form with Added Entry
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

' [4] Clear Userform Entries for Next
Me.U3a_Input_MUName.Text = ""
Me.U3a_Input_MUCO2Prod.Value = ""
Me.U3a_Input_MUCO2Cons.Value = ""
Me.U3a_Input_MUCost.Value = ""
End Sub


Private Sub U3a_Button_MUCancel_Click()
' [1] Re-Define Mass Utilities Range
Dim Interval_MU As String
Interval_MU = "DB_MUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_MU, RefersToLocal:="=OFFSET('B4'!$B$4,1,,COUNTA('B4'!$C:$C),2)"

Unload Me
End
End Sub


Private Sub U3a_Button_EUCancel_Click()
' [1] Re-Define Energy Utilities Range
Dim Interval_EU As String
Interval_EU = "DB_EUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_EU, RefersToLocal:="=OFFSET('B3'!$B$4,1,,COUNTA('B3'!$C:$C),2)"

Unload Me
End
End Sub


Private Sub U3a_ChooseExisting_EU_Click()
' [1] Show Userform
U3c_UtilityChooseExisting_EU.Show
End Sub


Private Sub U3a_ChooseExisting_MU_Click()
' [1] Show Userform
U3d_UtilityChooseExisting_MU.Show
End Sub
