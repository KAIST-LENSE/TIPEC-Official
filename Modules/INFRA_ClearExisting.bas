Attribute VB_Name = "INFRA_ClearExisting"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''   TIPEM CLEAR EXISTING MODULE   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ÀÌÁöÈ¯/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [INFRA] Clear Species Information
'''INACTIVE'''
Sub S1_ClearSpeciesInfo()
    ' [0] Are you sure?
    If MsgBox("This will erase the species information shown. Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' [1] Delete Species Info
    Worksheets("S1").Range("N15:O17").Select
    Selection.ClearContents
    Worksheets("S1").Range("N22").Select
    Selection.ClearContents
    Worksheets("S1").Range("N25").Select
    Selection.ClearContents
    Worksheets("S1").Range("N28").Select
    Selection.ClearContents
    Range("A1").Select
'''INACTIVE'''
End Sub


' [INFRA] Clear all Material Inventory Information
Sub S1_ClearMaterialsInfo()
    ' [0] Are you sure?
    If MsgBox("This will delete all materials for this project. Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' [1] Delete Materials in Backend
    Worksheets("B2").Range("B4:I2000").ClearContents
    
    ' [2] Scroll bar only visible if >= 20 Materials. If not, Sheet S1 table display range equals the respective range on B2 Datasheet
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
End Sub


' [INFRA] Clear all Utilities Information
Sub S2_ClearUtilitiesInfo()
    ' [0] Are you sure?
    If MsgBox("This will delete all currently specified utilities for this project. Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' [1] Delete Materials in Backend
    Worksheets("B3").Range("B5:F2000").ClearContents
    Worksheets("B4").Range("B5:F2000").ClearContents
    
    ' [2] Display EU or MU depending on display box color
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


' [INFRA] Clear all Transportation Information
Sub S2_ClearTransportInfo()
    ' [0] Are you sure?
    If MsgBox("This will delete all currently specified transportations for this project. Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' [1] Clear WORKSHEETS B11 Transportation Matrix
    TRANSPORT_Delete
    Worksheets("S2").Activate
    
    ' [2] Delete Materials in Backend
    Worksheets("B5").Range("B5:E2000").ClearContents
    
    ' [3] Display Transportations Data
    Set Transport_Data = Sheets("B5").Range("B5:E24")
    Set Transport_Display = Sheets("S2").Range("O15:R34")
    Transport_Display.Value = Transport_Data.Value
End Sub


' [INFRA] Clear Process Network
Sub S3_ClearNetwork()
    ' [0] Are you sure?
    If MsgBox("This will erase the network and all connections. Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' [1] Declare Variables
    Dim Shp As Shape
    Dim Shp2 As Shape
    Dim n_step As Integer
    Dim n_interval As Integer
    Dim n_mat As Integer
    Dim Current_Step As Integer
    Dim current_interval As Integer
    Dim max_interval As Integer
    Dim connection_p() As Integer
    Dim connection_s() As Integer
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim x As Integer
    
    ' [2] Determine Step and Interval Values
    n_step = Worksheets("S3").Range("H12").Value
    n_interval = Worksheets("S3").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value
    
    ' [3] Delete previous network figure
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
    For Each Shp2 In Worksheets("S8").Shapes
       If Not (Shp2.Type = msoOLEControlObject Or Shp2.Type = msoFormControl) Then Shp2.Delete
    Next Shp2
    Worksheets("S8").Range("C11").Value = ""

    ' [4] Delete Connectivity Data
    Worksheets("B8").Range("B4:F2000").ClearContents
    Worksheets("B9").Range("B4:F2000").ClearContents
    Worksheets("B11").Cells.Clear

    ' [5] Delete Connectivity Matrix
    Application.ScreenUpdating = False
    Application.Goto (Sheets("B7").Range("B4:CZ220"))
    Worksheets("B7").Range("B4:CZ220").ClearContents
    Worksheets("B12").Range("B4:CZ220").ClearContents
    Selection.Font.Bold = False
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Interior.TintAndShade = 0
    Worksheets("S7").Activate

    ' [6] Delete Process Interval Specification Table for Step 5
    TIPEM_Delete_IntervalSpecTable
    
    ' [7] Delete and Regenerate Transportation Matrixes
    TRANSPORT_Generate
    Worksheets("S8").Activate
    Application.ScreenUpdating = True
    
    ' [8] Delete Mass Balances Spreadsheet and Mass Balance Summary
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).ClearContents
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Font.Bold = False
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).UnMerge
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlInsideVertical).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlInsideHorizontal).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlEdgeTop).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlEdgeBottom).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlEdgeRight).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Borders(xlEdgeLeft).LineStyle = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Interior.Pattern = xlNone
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Interior.TintAndShade = 0
    Worksheets("O1").Range(Worksheets("O1").Cells(4, 2), Worksheets("O1").Cells(2 * n_mat, 11 * n_interval)).Interior.PatternTintAndShade = 0

    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).ClearContents
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Font.Bold = False
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).UnMerge
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlInsideVertical).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlInsideHorizontal).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlEdgeTop).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlEdgeBottom).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlEdgeRight).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Borders(xlEdgeLeft).LineStyle = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Interior.Pattern = xlNone
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Interior.TintAndShade = 0
    Worksheets("O2").Range(Worksheets("O2").Cells(4, 2), Worksheets("O2").Cells(220, 2 * n_mat)).Interior.PatternTintAndShade = 0

    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).ClearContents
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Font.Bold = False
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).UnMerge
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlInsideVertical).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlInsideHorizontal).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlEdgeTop).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlEdgeBottom).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlEdgeRight).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Borders(xlEdgeLeft).LineStyle = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Interior.Pattern = xlNone
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Interior.TintAndShade = 0
    Worksheets("O3").Range(Worksheets("O3").Cells(4, 2), Worksheets("O3").Cells(20 + n_interval, 12)).Interior.PatternTintAndShade = 0
    
    Worksheets("O4").Range("E6").ClearContents
    Worksheets("O4").Range("C7").ClearContents
    Worksheets("O4").Range("C10:C16").ClearContents
    Worksheets("O4").Range("C19:C23").ClearContents
    Worksheets("O4").Range("C26").ClearContents
    Worksheets("O4").Range("H7:H14").ClearContents
    Worksheets("O4").Range("C40:C42").ClearContents
    Worksheets("O4").Range("C52").ClearContents
    
    ' [9] Reset Checksums
    ' Mass Balance Available Checksum
    Worksheets("O1").Range("F2").Value = 0
    Worksheets("O2").Range("F2").Value = 0
    Worksheets("O3").Range("F2").Value = 0
    Worksheets("O4").Range("F2").Value = 0
    Worksheets("O4").Range("H2").Value = 0
End Sub



