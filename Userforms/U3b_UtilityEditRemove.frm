VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U3b_UtilityEditRemove 
   Caption         =   "Select a Utility to Edit or Remove"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   OleObjectBlob   =   "U3b_UtilityEditRemove.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U3b_UtilityEditRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U3b_EnergyUtil_List_Change()
On Error Resume Next
' [1] Display Selected Material Info
U3b_Display_EUtil.Text = U3b_EnergyUtil_List.Column(0) & "   |   " & U3b_EnergyUtil_List.Column(1)

' [2] Energy Utility Selected from List
Dim EUName As Double
If Me.U3b_EnergyUtil_List.Value = "" Then
MsgBox "You must specify an Energy Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Energy Utility Selected as Indexing String
EUName = U3b_EnergyUtil_List.Column(0)

' [4] Find the Row of the Indexed Energy Utility, then Load Values
Me.U3b_Input_EUCO2Prod.Value = Application.WorksheetFunction.VLookup(EUName, Sheets("B3").Range("B5:F2000"), 3, 0)
Me.U3b_Input_EUCO2Cons.Value = Application.WorksheetFunction.VLookup(EUName, Sheets("B3").Range("B5:F2000"), 4, 0)
Me.U3b_Input_EUCost.Value = Application.WorksheetFunction.VLookup(EUName, Sheets("B3").Range("B5:F2000"), 5, 0)
End Sub


Private Sub U3b_Button_RemoveEU_Click()
On Error Resume Next
' [1] Declare Variables
Dim msg As String
Dim ans As String
Dim rowselect As Double
Dim EUFound As Range

' [2] Must Select Energy Utility Before Proceeding
If Me.U3b_EnergyUtil_List.Value = "" Then
MsgBox "You must specify an Energy Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Energy Utility Name as String. Find Cell with String Value and Get Row Index
EUName2 = U3b_EnergyUtil_List.Column(1)
Set EUFound = Sheets("B3").Range("B5:F2000").Find(What:=EUName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Move Row Index After Energy Utility has been Deleted
If MsgBox("Are you sure you want to delete " & EUName2 & "?", vbYesNo) = vbNo Then Exit Sub
rowselect = EUFound.Row
Sheets("B3").Rows(rowselect).EntireRow.Delete
rowselect = rowselect - 1

' [5] Re-Define Energy Utilities Range
Dim Interval_EU As String
Interval_EU = "DB_EUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_EU, RefersToLocal:="=OFFSET('B3'!$B$4,1,,COUNTA('B3'!$C:$C),2)"

' [6] Auto-Re-Number the Material Name Index After Deleting Row
Dim Count As Integer
Dim NumbMat As Integer
Count = 1
NumbMat = Sheets("B3").Range("C1").Value
Application.EnableEvents = False
Do While Count <= NumbMat
    Sheets("B3").Range("B" & Count + 4).Value = Count
    Count = Count + 1
Loop
Application.EnableEvents = True

' [7] Display EU or MU depending on display box color
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
End Sub



Private Sub U3b_Button_UpdateEU_Click()
On Error Resume Next
' [1] Must select an energy utility to proceed updating
If Me.U3b_EnergyUtil_List.Value = "" Then
MsgBox "You must specify an Energy Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [2] Declare Variables
Dim EUName2 As Double
Dim EUFound As Range
Dim currentrow As Double

' [3] Energy Utility Name as String. Find Cell with String Value and Get Row Index
EUName2 = U3b_EnergyUtil_List.Column(0)
Set EUFound = Sheets("B3").Range("B5:B2000").Find(What:=EUName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Assign Row Index to a variable, current Row
currentrow = EUFound.Row

' [5] Update Data Entry
Sheets("B3").Cells(currentrow, 4) = U3b_Input_EUCO2Prod.Value
Sheets("B3").Cells(currentrow, 5) = U3b_Input_EUCO2Cons.Value
Sheets("B3").Cells(currentrow, 6) = U3b_Input_EUCost.Value

' [6] Re-Define Energy Utilities Range
Dim Interval_EU As String
Interval_EU = "DB_EUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_EU, RefersToLocal:="=OFFSET('B3'!$B$4,1,,COUNTA('B3'!$C:$C),2)"
    
' [7] Display EU or MU depending on display box color
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

' [8] Display Energy Utility Successfully Updated
ans = MsgBox(EUName2 & " has been successfully updated", vbOKOnly, "TIPEM- Energy Utility Updated")
If ans = vbYes Then
U3b_UtilityEditRemove.Show
End If
End Sub


Private Sub U3b_Button_Cancel1_Click()
Unload Me
End
End Sub


Private Sub U3b_MassUtil_List_Change()
On Error Resume Next
' [1] Display Selected Material Info
U3b_Display_MUtil.Text = U3b_MassUtil_List.Column(0) & "   |   " & U3b_MassUtil_List.Column(1)

' [2] Mass Utility Selected from List
Dim MUName As Double
If Me.U3b_MassUtil_List.Value = "" Then
MsgBox "You must specify a Mass Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Mass Utility Selected as Indexing String
MUName = U3b_MassUtil_List.Column(0)

' [4] Find the Row of the Indexed Mass Utility, then Load Values
Me.U3b_Input_MUCO2Prod.Value = Application.WorksheetFunction.VLookup(MUName, Sheets("B4").Range("B5:F2000"), 3, 0)
Me.U3b_Input_MUCO2Cons.Value = Application.WorksheetFunction.VLookup(MUName, Sheets("B4").Range("B5:F2000"), 4, 0)
Me.U3b_Input_MUCost.Value = Application.WorksheetFunction.VLookup(MUName, Sheets("B4").Range("B5:F2000"), 5, 0)
End Sub


Private Sub U3b_Button_RemoveMU_Click()
On Error Resume Next
' [1] Declare Variables
Dim msg As String
Dim ans As String
Dim rowselect As Double
Dim MUFound As Range

' [2] Must Select Mass Utility Before Proceeding
If Me.U3b_MassUtil_List.Value = "" Then
MsgBox "You must specify a Mass Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [3] Mass Utility Name as String. Find Cell with String Value and Get Row Index
MUName2 = U3b_MassUtil_List.Column(1)
Set MUFound = Sheets("B4").Range("B5:F2000").Find(What:=MUName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Move Row Index After Mass Utility has been Deleted
If MsgBox("Are you sure you want to delete " & MUName2 & "?", vbYesNo) = vbNo Then Exit Sub
rowselect = MUFound.Row
Sheets("B4").Rows(rowselect).EntireRow.Delete
rowselect = rowselect - 1

' [5] Re-Define Mass Utilities Range
Dim Interval_MU As String
Interval_MU = "DB_MUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_MU, RefersToLocal:="=OFFSET('B4'!$B$4,1,,COUNTA('B4'!$C:$C),2)"

' [6] Auto-Re-Number the Material Name Index After Deleting Row
Dim Count As Integer
Dim NumbMat As Integer
Count = 1
NumbMat = Sheets("B4").Range("C1").Value
Application.EnableEvents = False
Do While Count <= NumbMat
    Sheets("B4").Range("B" & Count + 4).Value = Count
    Count = Count + 1
Loop
Application.EnableEvents = True

' [7] Display EU or MU depending on display box color
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
End Sub


Private Sub U3b_Button_UpdateMU_Click()
On Error Resume Next
' [1] Must select a mass utility to proceed updating
If Me.U3b_MassUtil_List.Value = "" Then
MsgBox "You must specify a Mass Utility to be Edited or Removed!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [2] Declare Variables
Dim MUFound As Range
Dim currentrow As Double

' [3] Mass Utility Name as String. Find Cell with String Value and Get Row Index
MUName2 = U3b_MassUtil_List.Column(0)
Set MUFound = Sheets("B4").Range("B5:B2000").Find(What:=MUName2, LookIn:=xlFormulas, Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

' [4] Assign Row Index to a variable, current Row
currentrow = MUFound.Row

' [5] Update Data Entry
Sheets("B4").Cells(currentrow, 4) = U3b_Input_MUCO2Prod.Value
Sheets("B4").Cells(currentrow, 5) = U3b_Input_MUCO2Cons.Value
Sheets("B4").Cells(currentrow, 6) = U3b_Input_MUCost.Value

' [6] Re-Define Mass Utilities Range
Dim Interval_MU As String
Interval_MU = "DB_MUtil_List"
ActiveWorkbook.Names.Add Name:=Interval_MU, RefersToLocal:="=OFFSET('B4'!$B$4,1,,COUNTA('B4'!$C:$C),2)"

' [7] Display EU or MU depending on display box color
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

' [8] Display Material Successfully Updated
ans = MsgBox(MUName2 & " has been successfully updated", vbOKOnly, "TIPEM- Mass Utility Updated")
If ans = vbYes Then
    U3b_UtilityEditRemove.Show
End If
End Sub


Private Sub U3b_Button_Cancel2_Click()
Unload Me
End
End Sub

