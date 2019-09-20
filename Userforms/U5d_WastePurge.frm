VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5d_WastePurge 
   Caption         =   "WASTE PURGE STEP for Interval #-#"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   OleObjectBlob   =   "U5d_WastePurge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5d_WastePurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
On Error Resume Next
'[0] Simulate Button Click
    ActiveSheet.Shapes.Range(Array("Diamond 64")).ZOrder msoSendToBack


'[1] Find Current Interval
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim process_int As Integer
    Dim Num_Mat As Integer
    Dim Current_Row As Integer
    Dim Name_Row As Integer
    Dim Start_Index As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))
    
    ' Find the location of the current interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the name of the current interval and edit Userform
    Name_Row = Current_Row - 10 - process_int - 6 - process_int - 10 - 6 - Num_Int - (3 * (process_int + 6))
    U5d_WastePurge.Caption = "WASTE PURGE STEP for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5d_WS_Frame.Caption = "Waste Purge Specification for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5d_Utility_Frame.Caption = "Waste Purge Utility Consumption for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value


' [2] Create a Custom Range for ENERGY/MASS UTILITY with 2 Columns being EU Index and EU Name and 3rd Column being EU Consumption for that Index
    Dim EUtils As Integer
    Dim MUtils As Integer
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value
    
    Dim EU_array()
    Dim MU_array()
    ReDim EU_array(EUtils, 3)
    ReDim MU_array(MUtils, 3)

    For k = 1 To EUtils
        EU_array(k - 1, 0) = Worksheets("B3").Cells(4 + k, 2).Value
        EU_array(k - 1, 1) = Worksheets("B3").Cells(4 + k, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + k).Value = "" Then
            Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + k).Value = 0
            EU_array(k - 1, 2) = 0
        Else
            EU_array(k - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + k).Value
        End If
    Next k
    
    For m = 1 To MUtils
        MU_array(m - 1, 0) = Worksheets("B4").Cells(4 + m, 2).Value
        MU_array(m - 1, 1) = Worksheets("B4").Cells(4 + m, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + EUtils + m).Value = "" Then
            Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + EUtils + m).Value = 0
            MU_array(m - 1, 2) = 0
        Else
            MU_array(m - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + EUtils + m).Value
        End If
    Next m
    Me.U5d_EU_Combobox.list = EU_array
    Me.U5d_MU_Combobox.list = MU_array
    
    
'[3] Make a Range with 2 Columns being the Material Index, Material Name and append a 3rd Column corresponding to the Fraction Separated as Waste
    Dim WS_array()
    ReDim WS_array(Num_Mat, 3)

    For b = 1 To Num_Mat
        WS_array(b - 1, 0) = Worksheets("B2").Cells(3 + b, 2).Value
        WS_array(b - 1, 1) = Worksheets("B2").Cells(3 + b, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + b).Value = "" Then
            WS_array(b - 1, 2) = 0
        Else
            WS_array(b - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + b).Value
        End If
    Next b
    Me.U5d_Waste_Combobox.list = WS_array
End Sub




'***MATERIAL FOR WASTE SEP SELECTED: LOAD CURRENT WASTE SEPARATION VALUES (DEFAULT IS 0)***
Private Sub U5d_Waste_Combobox_Change()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5d_Waste_Combobox.Value = "" Then
        MsgBox "Please select a Material to be purged as Waste!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5d_WasteMaterial_Display = U5d_Waste_Combobox.Column(0) & "   |   " & U5d_Waste_Combobox.Column(1)


' [2] Find the current Waste Purge Fraction corresponding to the material selected
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim process_int As Integer
    Dim Num_Mat As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Index_Row As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5d_Waste_Combobox.Column(1) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' For the current interval, load the RM Specific Loading Value
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5d_WS_Fraction = 0
    Else
        U5d_WS_Fraction = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE SPECIFIED WASTE SEP FRACTION***
Private Sub U5d_Apply_Waste_Click()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5d_Waste_Combobox.Value = "" Then
        MsgBox "Please select a Material to be purged as Waste!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5d_WasteMaterial_Display = U5d_Waste_Combobox.Column(0) & "   |   " & U5d_Waste_Combobox.Column(1)


' [2] Find the current Waste Purge Fraction corresponding to the material selected
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim process_int As Integer
    Dim Num_Mat As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Index_Row As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5d_Waste_Combobox.Column(1) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' Save the entered Raw Material Loading in the respective cell
    If Me.U5d_WS_Fraction.Value < 0 Or Me.U5d_WS_Fraction.Value > 1 Then
        MsgBox "Waste Purge fraction must be between 0 and 1!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5d_WS_Fraction.Value


'[3] Make a Range with 2 Columns being the Material Index, Material Name and append a 3rd Column corresponding to the Fraction Separated as Waste
    Dim WS_array()
    ReDim WS_array(Num_Mat, 3)

    For g = 1 To Num_Mat
        WS_array(g - 1, 0) = Worksheets("B2").Cells(3 + g, 2).Value
        WS_array(g - 1, 1) = Worksheets("B2").Cells(3 + g, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + g).Value = "" Then
            WS_array(g - 1, 2) = 0
        Else
            WS_array(g - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + g).Value
        End If
    Next g
    Me.U5d_Waste_Combobox.list = WS_array
End Sub




'***SELECT AN EU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5d_EU_Combobox_Change()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5d_EU_Display.Text = U5d_EU_Combobox.Column(0) & "   |   " & U5d_EU_Combobox.Column(1)
    If Me.U5d_EU_Combobox.Value = "" Then
        MsgBox "Please select an Energy Utility before specifying its consumption!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If


' [2] Load Existing Utility Consumption Value
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim Num_Mat As Integer
    Dim process_int As Integer
    Dim Index_Row As Integer
    Dim EUtils As Integer
    Dim MUtils As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current ENERGY UTILITY selected
    For b = 1 To EUtils
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5d_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected EU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5d_EU_SpecificConsumption = 0
    Else
        U5d_EU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE ENERGY UTILITY CONSUMPTION VALUES***
Private Sub U5d_Apply_EU_Click()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5d_EU_Display.Text = U5d_EU_Combobox.Column(0) & "   |   " & U5d_EU_Combobox.Column(1)
    If Me.U5d_EU_Combobox.Value = "" Then
        MsgBox "Please select an Energy Utility before specifying its consumption!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If


' [2] Load Existing Utility Consumption Value
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim Num_Mat As Integer
    Dim process_int As Integer
    Dim Index_Row As Integer
    Dim EUtils As Integer
    Dim MUtils As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current ENERGY UTILITY selected
    For b = 1 To EUtils
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5d_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected EU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5d_EU_SpecificConsumption.Value


' [3] Reload Custom Range for Combobox to Update Values
    ' Create a Custom Range for ENERGY UTILITY with 2 Columns being EU Index and EU Name and 3rd Column being EU Consumption for that Index
    Dim EU_array()
    ReDim EU_array(EUtils, 3)

    For k = 1 To EUtils
        EU_array(k - 1, 0) = Worksheets("B3").Cells(4 + k, 2).Value
        EU_array(k - 1, 1) = Worksheets("B3").Cells(4 + k, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + k).Value = "" Then
            EU_array(k - 1, 2) = 0
        Else
            EU_array(k - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + k).Value
        End If
    Next k
    Me.U5d_EU_Combobox.list = EU_array
End Sub




'***SELECT AN MU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5d_MU_Combobox_Change()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5d_MU_Display.Text = U5d_MU_Combobox.Column(0) & "   |   " & U5d_MU_Combobox.Column(1)
    If Me.U5d_MU_Combobox.Value = "" Then
        MsgBox "Please select a Mass Utility before specifying its consumption!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If


' [2] Load Existing Utility Consumption Value
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim Num_Mat As Integer
    Dim process_int As Integer
    Dim Index_Row As Integer
    Dim EUtils As Integer
    Dim MUtils As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current MASS UTILITY selected
    For b = 1 To MUtils
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5d_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected MU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5d_MU_SpecificConsumption = 0
    Else
        U5d_MU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE MASS UTILITY CONSUMPTION VALUES***
Private Sub U5d_Apply_MU_Click()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5d_MU_Display.Text = U5d_MU_Combobox.Column(0) & "   |   " & U5d_MU_Combobox.Column(1)
    If Me.U5d_MU_Combobox.Value = "" Then
        MsgBox "Please select a Mass Utility before specifying its consumption!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If


' [2] Load Existing Utility Consumption Value
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim Current_Col As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim Num_Mat As Integer
    Dim process_int As Integer
    Dim Index_Row As Integer
    Dim EUtils As Integer
    Dim MUtils As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + (3 * (process_int + 6))
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current MASS UTILITY selected
    For b = 1 To MUtils
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5d_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected MU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5d_MU_SpecificConsumption.Value


' [3] Reload Custom Range for Combobox to Update Values
    ' Create a Custom Range for MASS UTILITY with 2 Columns being EU Index and EU Name and 3rd Column being EU Consumption for that Index
    Dim MU_array()
    ReDim MU_array(MUtils, 3)
    
    For m = 1 To MUtils
        MU_array(m - 1, 0) = Worksheets("B4").Cells(4 + m, 2).Value
        MU_array(m - 1, 1) = Worksheets("B4").Cells(4 + m, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + EUtils + m).Value = "" Then
            MU_array(m - 1, 2) = 0
        Else
            MU_array(m - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + Num_Mat + EUtils + m).Value
        End If
    Next m
    Me.U5d_MU_Combobox.list = MU_array
End Sub



'***CLOSE USERFORM***
Private Sub U5d_Close_Click()
Unload Me
End
End Sub

