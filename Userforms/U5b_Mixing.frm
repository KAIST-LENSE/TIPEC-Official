VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5b_Mixing 
   Caption         =   "MIXING STEP for Interval #-#"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   OleObjectBlob   =   "U5b_Mixing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5b_Mixing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

'***INITIALIZE USERFORM***
Private Sub UserForm_Initialize()
On Error Resume Next
'[0] Simulate Button Click
    ActiveSheet.Shapes.Range(Array("Oval 59")).ZOrder msoSendToBack


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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6
    
    ' Find the location of the current interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the name of the current interval and edit Userform
    Name_Row = Current_Row - Num_Int - 6 - 10 - process_int - 6
    U5b_Mixing.Caption = "MIXING STEP for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5b_RMFrame.Caption = "Raw Material Specification for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5b_Utility_Frame.Caption = "Mixing Utility Consumption for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value


'[2] Populate Combobox with Raw Material and RM Specific Loading Info
    Dim RM_array()
    ReDim RM_array(Num_Mat, 2)

    For b = 1 To Num_Mat
        RM_array(b - 1, 0) = Worksheets("B2").Cells(3 + b, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + b).Value = "" Then
            RM_array(b - 1, 1) = 0
        Else
            RM_array(b - 1, 1) = Worksheets("B10").Cells(Current_Row, 3 + b).Value
        End If
    Next b
    Me.U5b_RMList.list = RM_array


' [3] If Basis Material already specified, show it
    Dim Basis_Row As Integer
    Dim Basis_Index As Integer
    Dim Basis_Column As Integer
    Dim BasisMat_Ind As Integer
    
    Basis_Index = Start_Index - process_int - 6
    Basis_Row = Current_Row - process_int - 6
    
    For i = 1 To Num_Mat
        If Worksheets("B10").Cells(Basis_Row, 3 + i).Value = 1 Then
            Basis_Column = 3 + i
            BasisMat_Ind = i
            Me.U5b_BasisM_Display = BasisMat_Ind & "   |   " & Worksheets("B10").Cells(Basis_Index, Basis_Column).Value
        End If
    Next i


' [4] Create a Custom Range for ENERGY/MASS UTILITY with 2 Columns being EU Index and EU Name and 3rd Column being EU Consumption for that Index
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
    Me.U5b_EU_Combobox.list = EU_array
    Me.U5b_MU_Combobox.list = MU_array
End Sub




'***SELECT RM BASIS MATERIAL***
Private Sub U5b_RawMaterial_Combobox_Change()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5b_RawMaterial_Combobox.Value = "" Then
        MsgBox "Please select a Basis Material for Raw Material addition!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
    
' [2] Display Selected Material Info
    ' Resets currently visible RM Info
    U5b_RM_Loading_Display = ""
    U5b_BasisM_Display = U5b_RawMaterial_Combobox.Column(0) & "   |   " & U5b_RawMaterial_Combobox.Column(1)
    U5b_Label6 = "tons/ton-" & U5b_RawMaterial_Combobox.Column(1)
    
    
' [3] Find the Current Row in the Basis Material Matrix
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Num_Int As Integer
    Dim Num_Steps As Integer
    Dim process_int As Integer
    Dim Num_Mat As Integer
    Dim Current_Row As Integer
    Dim Index_Row As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Steps = Worksheets("S4").Range("H12").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    process_int = Num_Int - Worksheets("S4").Range("F13").Value - Worksheets("S4").Cells(14 + Num_Steps, 6).Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10
    
    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Assign 1 if matches the Basis Material Selected, 0 if Else
    For b = 1 To Num_Mat
        If U5b_RawMaterial_Combobox.Column(1) = Worksheets("B10").Cells(Start_Index, 3 + b).Value Then
            Worksheets("B10").Cells(Current_Row, 3 + b).Value = 1
        Else
            Worksheets("B10").Cells(Current_Row, 3 + b).Value = 0
        End If
    Next b
End Sub




'***SELECT A RAW MATERIAL TO SPECIFY ITS LOADING***
Private Sub U5b_RMList_Change()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5b_RMList.Value = "" Then
        MsgBox "Please select a Raw Material before entering its specific loading!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
    
' [2] Display Selected Material Info
    If Me.U5b_RawMaterial_Combobox = "" Then
        MsgBox "Please select a Basis Material first before specifying Raw Material addition!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5b_RM_Loading_Display = U5b_RMList.Column(0) & "   |   " & U5b_RMList.Column(1) & " tons/ton-" & U5b_RawMaterial_Combobox.Column(1)


' [3] Find the current specific loading of the raw material selected
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5b_RMList.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' For the current interval, load the RM Specific Loading Value
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5b_RM_Specification = 0
    Else
        U5b_RM_Specification = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE SPECIFIC RAW MATERIAL ADDITION***
Private Sub U5b_Apply_RawMaterial_Click()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5b_RMList.Value = "" Then
        MsgBox "Please select a Raw Material to be added!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
    
' [2] Find the current specific loading of the raw material selected
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5b_RMList.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' Save the entered Raw Material Loading in the respective cell
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5b_RM_Specification.Value


' [3] Update the ComboBox
    Dim RM_array()
    ReDim RM_array(Num_Mat, 2)

    For b = 1 To Num_Mat
        RM_array(b - 1, 0) = Worksheets("B2").Cells(3 + b, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + b).Value = "" Then
            RM_array(b - 1, 1) = 0
        Else
            RM_array(b - 1, 1) = Worksheets("B10").Cells(Current_Row, 3 + b).Value
        End If
    Next b
    Me.U5b_RMList.list = RM_array
End Sub




'***SELECT AN EU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5b_EU_Combobox_Change()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5b_EU_Display.Text = U5b_EU_Combobox.Column(0) & "   |   " & U5b_EU_Combobox.Column(1)
    If Me.U5b_EU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5b_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected EU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5b_EU_SpecificConsumption = 0
    Else
        U5b_EU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE ENERGY UTILITY CONSUMPTION VALUES***
Private Sub U5b_Apply_EU_Click()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5b_EU_Display.Text = U5b_EU_Combobox.Column(0) & "   |   " & U5b_EU_Combobox.Column(1)
    If Me.U5b_EU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5b_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected EU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5b_EU_SpecificConsumption.Value


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
    Me.U5b_EU_Combobox.list = EU_array
End Sub




'***SELECT AN MU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5b_MU_Combobox_Change()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5b_MU_Display.Text = U5b_MU_Combobox.Column(0) & "   |   " & U5b_MU_Combobox.Column(1)
    If Me.U5b_MU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5b_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected MU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5b_MU_SpecificConsumption = 0
    Else
        U5b_MU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE MASS UTILITY CONSUMPTION VALUES***
Private Sub U5b_Apply_MU_Click()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5b_MU_Display.Text = U5b_MU_Combobox.Column(0) & "   |   " & U5b_MU_Combobox.Column(1)
    If Me.U5b_MU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5b_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected MU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5b_MU_SpecificConsumption.Value


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
    Me.U5b_MU_Combobox.list = MU_array
End Sub




'***CLOSE USERFORM***
Private Sub U5b_Close_Click()
Unload Me
End
End Sub
