VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5c_Reaction 
   Caption         =   "REACTION STEP for Interval #-#"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   OleObjectBlob   =   "U5c_Reaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5c_Reaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***SELECT MATERIAL CORRESPONDING TO BIOMASS AND LIMITING NUTRIENT***
Private Sub U5c_CB_Biomass_Change()
' [1] Material Selected from List
If Me.U5c_CB_Biomass.Value = "" Then
MsgBox "Please select the Material representing Algal Biomass!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [1] Display Selected Material Info
U5c_Display_Biomass.Text = U5c_CB_Biomass.Column(0) & "   |   " & U5c_CB_Biomass.Column(1)
End Sub
Private Sub U5c_CB_LN_Change()
' [1] Material Selected from List
Lnn = Worksheets("B1").Range("C11").Value
If Me.U5c_CB_LN.Value = "" Then
MsgBox "Please select the Material corresponding to the Limiting Nutrient:" & Lnn, vbExclamation, "TIPEM- Error"
Exit Sub
End If

' [1] Display Selected Material Info
U5c_Display_LN.Text = U5c_CB_LN.Column(0) & "  |  " & U5c_CB_LN.Column(1)
End Sub




'***CULTIVATION SELECTED***
Private Sub U5c_Checkbox_Click()
If U5c_Checkbox.Value = True Then
    Me.U5c_CultivationFrame.Visible = True
    Me.U5c_Skip.Value = False
Else
    Me.U5c_CultivationFrame.Visible = False
    Me.U5c_Skip.Value = False
End If
End Sub
Private Sub U5c_Skip_Click()
    Me.U5c_CultivationFrame.Visible = False
    Me.U5c_Checkbox.Value = False
End Sub




'***USERFORM INITIALIZE***
Private Sub UserForm_Initialize()
On Error Resume Next
'[0] Simulate Button Click
    ActiveSheet.Shapes.Range(Array("Group 60")).ZOrder msoSendToBack


'[1] Disable Cultivation Specification Box by Default
    Me.U5c_CultivationFrame.Visible = False
    ' Display Selected Limiting Nutrient in S1 in Textbox
    U5c_Label_LN.Caption = Worksheets("B1").Range("C11").Value

'[2] Find Current Interval
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10
    
    ' Find the location of the current interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the name of the current interval and edit Userform
    Name_Row = Current_Row - 10 - process_int - 6 - process_int - 10 - 6 - Num_Int
    U5c_Reaction.Caption = "REACTION STEP for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5c_Reaction_Frame.Caption = "Reaction Specification for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value
    U5c_Utility_Frame.Caption = "Reaction Utility Consumption for Interval [" & Worksheets("B10").Cells(Name_Row, 2).Value & "-" & Worksheets("B10").Cells(Name_Row, 3).Value & "] " & Worksheets("B10").Cells(Name_Row, 4).Value


' [3] Create a Custom Range for ENERGY/MASS UTILITY with 2 Columns being EU Index and EU Name and 3rd Column being EU Consumption for that Index
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
    Me.U5c_EU_Combobox.list = EU_array
    Me.U5c_MU_Combobox.list = MU_array


' [4] Create a Custom Range for Key Reactant Fractional Conversion Specificiation
    Dim KR_array()
    ReDim KR_array(Num_Mat, 3)

    For b = 1 To Num_Mat
        KR_array(b - 1, 0) = Worksheets("B2").Cells(3 + b, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + b).Value = 0 Or Worksheets("B10").Cells(Current_Row, 3 + b).Value = "" Then
            KR_array(b - 1, 2) = 0
            KR_array(b - 1, 1) = ""
        Else
            KR_array(b - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + b).Value
            KR_array(b - 1, 1) = "Key Reactant"
        End If
    Next b
    Me.U5c_KeyReactant_Combobox.list = KR_array


' [5] Create a Custom Range for Non-Key Reactant Fractional Conversion Specificiation
    Dim NKR_array()
    ReDim NKR_array(Num_Mat, 2)
    Dim NKR_Current_Row As Integer
    NKR_Current_Row = Current_Row + process_int + 6
    
    For j = 1 To Num_Mat
        NKR_array(j - 1, 0) = Worksheets("B2").Cells(3 + j, 3).Value
        If Worksheets("B10").Cells(NKR_Current_Row, 3 + j).Value = 0 Or Worksheets("B10").Cells(NKR_Current_Row, 3 + j).Value = "" Then
            Worksheets("B10").Cells(NKR_Current_Row, 3 + j).Value = 0
            NKR_array(j - 1, 1) = 0
        Else
            NKR_array(j - 1, 1) = Worksheets("B10").Cells(NKR_Current_Row, 3 + j).Value
        End If
    Next j
    Me.U5c_NonKeyReactant_Combobox.list = NKR_array


' [6] If Key Material already Specified, Show It
    For i = 1 To Num_Mat
        If Worksheets("B10").Cells(Current_Row, 3 + i).Value <> 0 Then
            Basis_Column = 3 + i
            Me.U5c_KR_Display = Worksheets("B10").Cells(Start_Index, Basis_Column) & "   |    Key Reactant"
        End If
    Next i
    U5c_NKR_Label3.Caption = "ton/ton-" & Worksheets("B10").Cells(Start_Index, Basis_Column).Value


' [7] Create a Custom Range for Product Mass Fractional Yield
    Dim PROD_array()
    ReDim PROD_array(Num_Mat, 2)
    Dim PROD_Current_Row As Integer
    PROD_Current_Row = Current_Row + (2 * (process_int + 6))
    
    For aa = 1 To Num_Mat
        PROD_array(aa - 1, 0) = Worksheets("B2").Cells(3 + aa, 3).Value
        If Worksheets("B10").Cells(PROD_Current_Row, 3 + aa).Value = 0 Or Worksheets("B10").Cells(PROD_Current_Row, 3 + aa).Value = "" Then
            Worksheets("B10").Cells(PROD_Current_Row, 3 + aa).Value = 0
            PROD_array(aa - 1, 1) = 0
        Else
            PROD_array(aa - 1, 1) = Worksheets("B10").Cells(PROD_Current_Row, 3 + aa).Value
        End If
    Next aa
    Me.U5c_Prod_Combobox.list = PROD_array


' [8] Check that the sum of Product Fractional Yields equal 1
    Dim FracYield_Sum As Double
    FracYield_Sum = 0
    
    For cc = 1 To Num_Mat
        FracYield_Sum = FracYield_Sum + Worksheets("B10").Cells(PROD_Current_Row, 3 + cc).Value
    Next cc
    U5c_ProdYield_Sum = FracYield_Sum
End Sub




'***SELECT A MATERIAL AS A KEY REACTANT***
Private Sub U5c_KeyReactant_Combobox_Change()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5c_KeyReactant_Combobox.Value = "" Then
        MsgBox "Please specify a material as a Key Reactant!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_KR_Display = U5c_KeyReactant_Combobox.Column(0) & "   |   " & U5c_KeyReactant_Combobox.Column(1)


' [2] Find the current interval and load Key Reactant Fractional Conversion Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_KeyReactant_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' For the current interval, load the RM Specific Loading Value
    U5c_KeyReactant_FC = Worksheets("B10").Cells(Current_Row, Current_Col).Value
End Sub




'***SAVE SPECIFIED KEY REACTANT INFO***
Private Sub U5c_KR_Save_Click()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5c_KeyReactant_Combobox.Value = "" Then
        MsgBox "Please specify a material as a Key Reactant!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_KR_Display = U5c_KeyReactant_Combobox.Column(0) & "   |   " & U5c_KeyReactant_Combobox.Column(1)


' [2] Find the current interval and load Key Reactant Fractional Conversion Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_KeyReactant_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' Reset the Matrix of Fractional Conversions
    For i = 1 To Num_Mat
        Worksheets("B10").Cells(Current_Row, 3 + i).Value = 0
    Next i
    
    ' Save the entered Key Reactant Frac Conv Loading in the respective cell
    If Me.U5c_KeyReactant_FC.Value < 0 Or Me.U5c_KeyReactant_FC.Value > 1 Then
        MsgBox "Fractional Conversion must be between 0 and 1!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5c_KeyReactant_FC.Value


' [3] Create a Custom Range for Key Reactant Fractional Conversion Specificiation
    Dim KR_array()
    ReDim KR_array(Num_Mat, 3)

    For b = 1 To Num_Mat
        KR_array(b - 1, 0) = Worksheets("B2").Cells(3 + b, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + b).Value = 0 Then
            KR_array(b - 1, 1) = ""
        Else
            KR_array(b - 1, 2) = Worksheets("B10").Cells(Current_Row, 3 + b).Value
            KR_array(b - 1, 1) = "Key Reactant"
        End If
    Next b
    Me.U5c_KeyReactant_Combobox.list = KR_array


' [4] Update Caption on Page 2 since Key Material Specified
    U5c_NKR_Label3.Caption = "ton/ton-" & U5c_KeyReactant_Combobox.Column(0)
End Sub




'***SELECT A MATERIAL AS A NON-KEY REACTANT***
Private Sub U5c_NonKeyReactant_Combobox_Change()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5c_NonKeyReactant_Combobox.Value = "" Then
        MsgBox "Please specify a material as a Non-Key Reactant!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_NKR_Display = U5c_NonKeyReactant_Combobox.Column(0) & "   |   " & U5c_NonKeyReactant_Combobox.Column(1)

' [2] Find the current interval and load Non-Key Reactant Specific Consumption Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_NonKeyReactant_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' For the current interval, load the RM Specific Loading Value
    U5c_NKR_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
End Sub




'***SAVE SPECIFIED NON-KEY REACTANT CONSUMPTION***
Private Sub U5c_NKR_Apply_Click()
On Error Resume Next
' [1] Material Selected from List
    If Me.U5c_NonKeyReactant_Combobox.Value = "" Then
        MsgBox "Please specify a material as a Non-Key Reactant!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_NKR_Display = U5c_NonKeyReactant_Combobox.Column(0)


' [2] Find the current interval and load Key Reactant Fractional Conversion Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_NonKeyReactant_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b

    ' Save the entered NonKey Reactant Specific Consumption in the respective cell
    If Me.U5c_NKR_SpecificConsumption.Value < 0 Then
        MsgBox "Specific Consumption must be a positive value!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5c_NKR_SpecificConsumption.Value


' [3] Create a Custom Range for Key Reactant Fractional Conversion Specificiation
    Dim NKR_array()
    ReDim NKR_array(Num_Mat, 2)
    
    For j = 1 To Num_Mat
        NKR_array(j - 1, 0) = Worksheets("B2").Cells(3 + j, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + j).Value = 0 Or Worksheets("B10").Cells(Current_Row, 3 + j).Value = "" Then
            Worksheets("B10").Cells(Current_Row, 3 + j).Value = 0
            NKR_array(j - 1, 1) = 0
        Else
            NKR_array(j - 1, 1) = Worksheets("B10").Cells(Current_Row, 3 + j).Value
        End If
    Next j
    Me.U5c_NonKeyReactant_Combobox.list = NKR_array
End Sub




'***SELECT A PRODUCT***
Private Sub U5c_Prod_Combobox_Change()
On Error Resume Next
' [1] Product Selected from List
    If Me.U5c_Prod_Combobox.Value = "" Then
        MsgBox "Please specify a material that is produced from the reaction!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_Prod_Display = U5c_Prod_Combobox.Column(0) & "   |   " & U5c_Prod_Combobox.Column(1)

' [2] Find the current interval and load Product Fractional Yield Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + process_int + 6 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_Prod_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b
    
    ' For the current interval, load the RM Specific Loading Value
    U5c_Prod_FracYield = Worksheets("B10").Cells(Current_Row, Current_Col).Value
End Sub




'***SAVE SPECIFIED PRODUCT FRACTIONAL MASS YIELD***
Private Sub U5c_Prod_Apply_Click()
On Error Resume Next
' [1] Product Selected from List
    If Me.U5c_Prod_Combobox.Value = "" Then
        MsgBox "Please specify a material that is produced from the reaction!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U5c_Prod_Display = U5c_Prod_Combobox.Column(0) & "   |   " & U5c_Prod_Combobox.Column(1)


' [2] Find the current interval and load Key Reactant Fractional Conversion Already Specified
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10 + process_int + 6 + process_int + 6

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current material specified
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(Start_Index, 3 + b).Value = U5c_Prod_Combobox.Column(0) Then
            Current_Col = 3 + b
        End If
    Next b

    ' Save the entered Product Fractional Mass Yield in the Cell
    If Me.U5c_Prod_FracYield.Value < 0 Then
        MsgBox "Product Fractional Yield must be a positive value!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5c_Prod_FracYield.Value
    

' [3] Check that the sum of Product Fractional Yields equal 1
    Dim FracYield_Sum As Double
    Dim React_Sum As Double
    Dim React_Row As Integer
    
    React_Sum = 0
    FracYield_Sum = 0
    React_Row = Current_Row - 6 - process_int - 6 - process_int
    
    For bb = 1 To Num_Mat
        React_Sum = React_Sum + Worksheets("B10").Cells(React_Row, 3 + bb).Value
    Next bb
    
    If React_Sum <> 0 Then
        For cc = 1 To Num_Mat
            FracYield_Sum = FracYield_Sum + Worksheets("B10").Cells(Current_Row, 3 + cc).Value
        Next cc
        U5c_ProdYield_Sum = FracYield_Sum
    End If


' [4] Create a Custom Range for Product Mass Fractional Yield
    Dim PROD_array()
    ReDim PROD_array(Num_Mat, 2)
    
    For aa = 1 To Num_Mat
        PROD_array(aa - 1, 0) = Worksheets("B2").Cells(3 + aa, 3).Value
        If Worksheets("B10").Cells(Current_Row, 3 + aa).Value = 0 Or Worksheets("B10").Cells(Current_Row, 3 + aa).Value = "" Then
            Worksheets("B10").Cells(Current_Row, 3 + aa).Value = 0
            PROD_array(aa - 1, 1) = 0
        Else
            PROD_array(aa - 1, 1) = Worksheets("B10").Cells(Current_Row, 3 + aa).Value
        End If
    Next aa
    Me.U5c_Prod_Combobox.list = PROD_array
End Sub




'***SELECT AN EU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5c_EU_Combobox_Change()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5c_EU_Display.Text = U5c_EU_Combobox.Column(0) & "   |   " & U5c_EU_Combobox.Column(1)
    If Me.U5c_EU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10
    EUtils = Worksheets("B3").Range("C1").Value
    MUtils = Worksheets("B4").Range("C1").Value

    ' Find the row index of current process interval
    For a = 1 To process_int
        If Worksheets("B10").Cells(Start_Index + a, 2).Value = Current_Step And Worksheets("B10").Cells(Start_Index + a, 3).Value = Current_Int Then
            Current_Row = Start_Index + a
        End If
    Next a
    
    ' Find the column index of the current MASS UTILITY selected
    For b = 1 To EUtils
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5c_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected MU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5c_EU_SpecificConsumption = 0
    Else
        U5c_EU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE ENERGY UTILITY CONSUMPTION VALUES***
Private Sub U5c_Apply_EU_Click()
On Error Resume Next
' [1] Display Selected EU in Textbox
    U5c_EU_Display.Text = U5c_EU_Combobox.Column(0) & "   |   " & U5c_EU_Combobox.Column(1)
    If Me.U5c_EU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + b).Value = U5c_EU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected EU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5c_EU_SpecificConsumption.Value


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
    Me.U5c_EU_Combobox.list = EU_array
End Sub




'***SELECT AN MU, DISPLAY VALUE IF ALREADY SPECIFIED***
Private Sub U5c_MU_Combobox_Change()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5c_MU_Display.Text = U5c_MU_Combobox.Column(0) & "   |   " & U5c_MU_Combobox.Column(1)
    If Me.U5c_MU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5c_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, load the SPECIFIC CONSUMPTION value for the Selected MU
    If Worksheets("B10").Cells(Current_Row, Current_Col).Value = "" Then
        U5c_MU_SpecificConsumption = 0
    Else
        U5c_MU_SpecificConsumption = Worksheets("B10").Cells(Current_Row, Current_Col).Value
    End If
End Sub




'***SAVE MASS UTILITY CONSUMPTION VALUES***
Private Sub U5c_Apply_MU_Click()
On Error Resume Next
' [1] Display Selected MU in Textbox
    U5c_MU_Display.Text = U5c_MU_Combobox.Column(0) & "   |   " & U5c_MU_Combobox.Column(1)
    If Me.U5c_MU_Combobox.Value = "" Then
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
    Start_Index = 7 + Num_Int + 6 + Worksheets("S4").Range("F13").Value + 10 + process_int + 6 + process_int + 10
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
        If Worksheets("B10").Cells(Start_Index, 3 + Num_Mat + EUtils + b).Value = U5c_MU_Combobox.Column(1) Then
            Current_Col = 3 + Num_Mat + EUtils + b
        End If
    Next b
    
    ' For the current interval, save the SPECIFIC CONSUMPTION value for the Selected MU
    Worksheets("B10").Cells(Current_Row, Current_Col).Value = U5c_MU_SpecificConsumption.Value


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
    Me.U5c_MU_Combobox.list = MU_array
End Sub




'***CLOSE USERFORM***
Private Sub U5c_Close_Click()
Unload Me
End
End Sub
