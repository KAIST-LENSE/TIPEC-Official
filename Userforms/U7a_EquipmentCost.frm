VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U7a_EquipmentCost 
   Caption         =   "Equipment Cost Specification"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9510
   OleObjectBlob   =   "U7a_EquipmentCost.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U7a_EquipmentCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** UPON INITIALIZING, GENERABLE TEA TABLE IN O3 IF DATA NOT THERE ****
Private Sub UserForm_Initialize()
' [1] Declare Variables
    Dim a As Integer
    Dim b As Integer
    Dim head As Range
    Dim cell_selected As Range
    Dim Num_Steps As Integer
    Dim process_step As Integer
    Dim process_int As Integer
    Num_Steps = Worksheets("S4").Range("H12").Value + 2

' [2] Generate TEA Tables if O3 is Empty
    If Worksheets("O3").Range("B4").Value = "" Then
        '(A) Initialize TEA Tables
        Set head = Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Offset(1)
        With head
           .Value = "[1] Equipment Cost Tables"
           .Font.Bold = True
           .HorizontalAlignment = xlLeft
        End With
        '(B) Generate Table Headers
        Set cell_selected = Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Offset(1)
        With Range(cell_selected, cell_selected.Offset(, 1))
           .Merge
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .Value = "Index"
           .Font.Bold = True
           .Interior.Color = RGB(221, 235, 247)
        End With
        With cell_selected.Offset(1)
           .Value = "Process Step"
           .HorizontalAlignment = xlCenter
           .Font.Bold = True
           .Interior.Color = RGB(221, 235, 247)
        End With
        Set cell_selected = Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
        With cell_selected
           .Value = "Interval"
           .HorizontalAlignment = xlCenter
           .Font.Bold = True
           .Interior.Color = RGB(221, 235, 247)
        End With
        '(C) Equipment Cost Parameters and Entries
        Set cell_selected = Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
        With Range(cell_selected, cell_selected.Offset(, 5 - 1))
           .Merge
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .Value = "Equipment Cost Input Parameters"
           .Font.Bold = True
           .Interior.Color = RGB(221, 235, 247)
        End With
        With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(, 1)
         .Value = "Alpha ¥ð_kk^A, ($)"
         .Interior.Color = RGB(221, 235, 247)
        End With
        With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(, 2)
         .Value = "Beta ¥ð_kk^B, (tons)"
         .Interior.Color = RGB(221, 235, 247)
        End With
        With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(, 3)
         .Value = "Factor ¥ð_kk^C"
         .Interior.Color = RGB(221, 235, 247)
        End With
        With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(, 4)
         .Value = "Scaling Mass Capacity (tons)"
         .Interior.Color = RGB(221, 235, 247)
        End With
        With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(, 5)
         .Value = "Purchased Equipment Cost($)"
         .Interior.Color = RGB(221, 235, 247)
        End With
        '(D) Number Process Intervals
        For a = 1 To Num_Steps - 2
           process_step = Worksheets("S4").Range("E" & 13 + a).Value
           process_int = Worksheets("S4").Range("F" & 13 + a).Value
           For b = 1 To process_int
              With Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Offset(1)
                .Value = process_step
                .Interior.Color = RGB(221, 235, 247)
              End With
              With Worksheets("O3").Cells(Rows.Count, "C").End(xlUp).Offset(1)
                .Value = b
                .Interior.Color = RGB(221, 235, 247)
              End With
           Next b
        Next a
       '(E) Draw Cell Boundaries
        ColInd = Worksheets("O3").Cells(6, Columns.Count).End(xlToLeft).Column
        RowInd = Worksheets("O3").Cells(Rows.Count, "B").End(xlUp).Row
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
        Worksheets("O3").Range(Worksheets("O3").Cells(5, 2), Worksheets("O3").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    End If

' [3] Hide Specification Frames Until Checkbox Selected
    Me.U7a_Frame1.Visible = False
    Me.U7a_Frame2.Visible = False

' [4] Create a Custom Array of Process Intervals to be Displayed on Userform
    Dim total_intervals As Integer
    Dim n_feed As Integer
    Dim n_prod As Integer
    Dim n_proc As Integer
    
    total_intervals = Worksheets("S4").Range("H14").Value
    n_feed = Worksheets("S4").Range("F13").Value
    n_prod = Worksheets("S4").Cells(Rows.Count, "F").End(xlUp).Value
    n_proc = total_intervals - n_feed - n_prod
    
    Dim PI_Array()
    ReDim PI_Array(n_proc, 4)
    
    For a = 1 To n_proc
        PI_Array(a - 1, 0) = Worksheets("B7").Cells(7 + n_feed + a, 2).Value
        PI_Array(a - 1, 1) = Worksheets("B7").Cells(7 + n_feed + a, 3).Value
        PI_Array(a - 1, 2) = Worksheets("B10").Cells(7 + n_feed + a, 4).Value
        PI_Array(a - 1, 3) = Format(Worksheets("O3").Cells(6 + a, 8).Value, "$0,000.00")
    Next a
    Me.U7a_IntervalSelect_Combobox.list = PI_Array

' [5] Custom Array for Choosing Scaling Mass Capacity
    Dim Scaling_Array()
    ReDim Scaling_Array(4, 2)
    
    Scaling_Array(0, 0) = 1
    Scaling_Array(0, 1) = "Total Interval Inlet Mass"
    Scaling_Array(1, 0) = 2
    Scaling_Array(1, 1) = "Total Post-RM Mixing Mass"
    Scaling_Array(2, 0) = 3
    Scaling_Array(2, 1) = "Total Post-Reaction Mass"
    Scaling_Array(3, 0) = 4
    Scaling_Array(3, 1) = "Total Post-Waste Purge Mass"
    Me.U7a_Combobox1.list = Scaling_Array
End Sub





'**** SELECT A PROCESS INTERVAL TO ENTER THE EQUIPMENT COSTS *****
Private Sub U7a_IntervalSelect_Combobox_Change()
On Error Resume Next
' [1] Process Interval Selected from List
    If Me.U7a_IntervalSelect_Combobox.Value = "" Then
        MsgBox "Please select a Process Interval from the list!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U7a_IntervalSelect_Combobox.Column(3) = "" Then
        U7a_Interval_Display = "Step [" & U7a_IntervalSelect_Combobox.Column(0) & "-" & U7a_IntervalSelect_Combobox.Column(1) & "]         |         " & U7a_IntervalSelect_Combobox.Column(2) & "                             --- $"
    Else
        U7a_Interval_Display = "Step [" & U7a_IntervalSelect_Combobox.Column(0) & "-" & U7a_IntervalSelect_Combobox.Column(1) & "]         |         " & U7a_IntervalSelect_Combobox.Column(2) & "                           " & U7a_IntervalSelect_Combobox.Column(3)
    End If
    Me.U7a_Label8.Caption = "Enter a Vendor Cost Quote ($)" & vbNewLine & "for" & U7a_IntervalSelect_Combobox.Column(2)


' [2] Load Equipment Cost Parameter Values
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim total_intervals As Integer
    Dim n_feed As Integer
    Dim n_prod As Integer
    Dim n_proc As Integer
    
    Current_Step = Me.U7a_IntervalSelect_Combobox.Column(0)
    Current_Int = Me.U7a_IntervalSelect_Combobox.Column(1)
    total_intervals = Worksheets("S4").Range("H14").Value
    n_feed = Worksheets("S4").Range("F13").Value
    n_prod = Worksheets("S4").Cells(Rows.Count, "F").End(xlUp).Value
    n_proc = total_intervals - n_feed - n_prod
    
    ' Find the row index of current process interval
    For a = 1 To n_proc
        If Worksheets("O3").Cells(6 + a, 2).Value = Current_Step And Worksheets("O3").Cells(6 + a, 3).Value = Current_Int Then
            Current_Row = 6 + a
        End If
    Next a
    
    ' For the current interval, load the existing values
    If Worksheets("O3").Cells(Current_Row, 4).Value = 0 Then
        Me.U7a_RefA.Value = 0
    Else
        Me.U7a_RefA.Value = Worksheets("O3").Cells(Current_Row, 4).Value
    End If
    If Worksheets("O3").Cells(Current_Row, 5).Value = 0 Then
        Me.U7a_RefB.Value = 0
    Else
        Me.U7a_RefB.Value = Worksheets("O3").Cells(Current_Row, 5).Value
    End If
    If Worksheets("O3").Cells(Current_Row, 6).Value = 0 Then
        Me.U7a_Ref3.Value = 0
    Else
        Me.U7a_Ref3.Value = Worksheets("O3").Cells(Current_Row, 6).Value
    End If
    If Worksheets("O3").Cells(Current_Row, 8).Value = 0 Then
        Me.U7a_Vendor.Value = 0
    Else
        Me.U7a_Vendor.Value = Worksheets("O3").Cells(Current_Row, 8).Value
    End If
End Sub





'**** SELECT A SPECIFICATION METHOD
Private Sub U7a_Checkbox1_Click()
On Error Resume Next
' [1] Make sure a Process Interval is Selected
    If Me.U7a_IntervalSelect_Combobox.Value = "" Then
        MsgBox "Please select a Process Interval from the list!!", vbExclamation, "TIPEM- Error"
        U7a_Checkbox1.Value = False
        Exit Sub
    End If
    
' [2] If Checkbox is Selected, show the corresponding Frame
    If U7a_Checkbox1.Value = True Then
        Me.U7a_Frame1.Visible = True
        If U7a_Checkbox2.Value = True Then
            U7a_Checkbox2.Value = False
        End If
    Else
        Me.U7a_Frame1.Visible = False
    End If
End Sub
Private Sub U7a_Checkbox2_Click()
On Error Resume Next
' [1] Make sure a Process Interval is Selected
    If Me.U7a_IntervalSelect_Combobox.Value = "" Then
        MsgBox "Please select a Process Interval from the list!!", vbExclamation, "TIPEM- Error"
        U7a_Checkbox2.Value = False
        Exit Sub
    End If
    
' [2] If Checkbox is Selected, show the corresponding Frame
    If U7a_Checkbox2.Value = True Then
        Me.U7a_Frame2.Visible = True
        If U7a_Checkbox1.Value = True Then
            U7a_Checkbox1.Value = False
        End If
    Else
        Me.U7a_Frame2.Visible = False
    End If
End Sub
Private Sub U7a_Combobox1_Change()
    If Me.U7a_Combobox1.Value = "" Then
        MsgBox "Please choose a Scaling Mass!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    U7a_ScalingDisplay = U7a_Combobox1.Column(1)
End Sub





'**** APPLY EQUIPMENT COST PARAMETER *****
Private Sub U7a_Apply1_Click()
'On Error Resume Next
' [1] Make sure a Process Interval has been Selected
    If Me.U7a_IntervalSelect_Combobox.Value = "" Then
        MsgBox "Please select a Process Interval from the list!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U7a_Combobox1.Value = "" Then
        MsgBox "Please choose a Scaling Mass!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Find the corresponding row and enter values
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim This_Row As Integer
    Dim total_intervals As Integer
    Dim n_feed As Integer
    Dim n_prod As Integer
    Dim n_proc As Integer
    Current_Step = Me.U7a_IntervalSelect_Combobox.Column(0)
    Current_Int = Me.U7a_IntervalSelect_Combobox.Column(1)
    total_intervals = Worksheets("S4").Range("H14").Value
    n_feed = Worksheets("S4").Range("F13").Value
    n_prod = Worksheets("S4").Cells(Rows.Count, "F").End(xlUp).Value
    n_proc = total_intervals - n_feed - n_prod
    
    'Find the row index of current process interval
    For a = 1 To n_proc
        If Worksheets("O3").Cells(6 + a, 2).Value = Current_Step And Worksheets("O3").Cells(6 + a, 3).Value = Current_Int Then
            Current_Row = 6 + a
            This_Row = a
        End If
    Next a
    
    'Save current values to worksheet
    Starting_Index = 26 + n_prod + This_Row + Worksheets("B2").Range("K3").Value + Worksheets("B3").Range("C1").Value + Worksheets("B4").Range("C1").Value + Worksheets("B5").Range("C1").Value
    If U7a_Checkbox1.Value = True Then
        Worksheets("O3").Cells(Current_Row, 4).Value = Me.U7a_RefA.Value
        Worksheets("O3").Cells(Current_Row, 5).Value = Me.U7a_RefB.Value
        Worksheets("O3").Cells(Current_Row, 6).Value = Me.U7a_Ref3.Value
        Worksheets("O3").Cells(Current_Row, 7).Value = Worksheets("O2").Cells(Starting_Index, 3 + U7a_Combobox1.Column(0)).Value
    End If

' [3] Calculate the Equipment Cost and Enter the Value of the Equipment Cost in the corresponding Cell
    Worksheets("O3").Cells(Current_Row, 8).Value = Me.U7a_RefA.Value * ((Worksheets("O3").Cells(Current_Row, 7).Value / Me.U7a_RefB.Value) ^ Me.U7a_Ref3.Value)

' [4] Update the Array for Combobox Display
    Dim PI_Array()
    ReDim PI_Array(n_proc, 4)
    
    For a = 1 To n_proc
        PI_Array(a - 1, 0) = Worksheets("B7").Cells(7 + n_feed + a, 2).Value
        PI_Array(a - 1, 1) = Worksheets("B7").Cells(7 + n_feed + a, 3).Value
        PI_Array(a - 1, 2) = Worksheets("B10").Cells(7 + n_feed + a, 4).Value
        PI_Array(a - 1, 3) = Format(Worksheets("O3").Cells(6 + a, 8).Value, "$0,000.00")
    Next a
    Me.U7a_IntervalSelect_Combobox.list = PI_Array
    
' [5] Check if Equipment Costs for all Intervals have been Specified, if so, then update the CHECKSUM
    Worksheets("O3").Range("F2").Value = 1
    For a = 1 To n_proc
        If Worksheets("O3").Cells(6 + a, 8).Value = "" Then
            Worksheets("O3").Range("F2").Value = 0
        End If
    Next a
    If Worksheets("O3").Range("F2").Value = 1 Then
        Worksheets("S7").Shapes.Range(Array("TextBox 21")).Select
        With Selection
            .Text = "¡î  -  Equipment Costs for all intervals have been specified."
            .Font.Italic = msoTrue
            .Font.Size = 12
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
        End With
    Else
        Worksheets("S7").Shapes.Range(Array("TextBox 21")).Select
        With Selection
            .Text = "X  -  Equipment Costs for all intervals have been specified."
            .Font.Italic = msoTrue
            .Font.Size = 12
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(118, 97, 66)
        .Transparency = 0
        .Solid
        End With
    End If
    Range("A1").Select
    
' [6] Equipment Cost Calculated!!
    MsgBox "Purchased Equipment Cost Calculated!!", vbExclamation, "TIPEM- Notice"
End Sub
Private Sub U7a_Apply2_Click()
'On Error Resume Next
' [1] Make sure a Process Interval has been Selected
    If Me.U7a_IntervalSelect_Combobox.Value = "" Then
        MsgBox "Please select a Process Interval from the list!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Find the corresponding row and enter values
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Current_Row As Integer
    Dim total_intervals As Integer
    Dim n_feed As Integer
    Dim n_prod As Integer
    Dim n_proc As Integer
    Current_Step = Me.U7a_IntervalSelect_Combobox.Column(0)
    Current_Int = Me.U7a_IntervalSelect_Combobox.Column(1)
    total_intervals = Worksheets("S4").Range("H14").Value
    n_feed = Worksheets("S4").Range("F13").Value
    n_prod = Worksheets("S4").Cells(Rows.Count, "F").End(xlUp).Value
    n_proc = total_intervals - n_feed - n_prod
    
    'Find the row index of current process interval
    For a = 1 To n_proc
        If Worksheets("O3").Cells(6 + a, 2).Value = Current_Step And Worksheets("O3").Cells(6 + a, 3).Value = Current_Int Then
            Current_Row = 6 + a
        End If
    Next a

' [3] Calculate the Equipment Cost and Enter the Value of the Equipment Cost in the corresponding Cell
    Worksheets("O3").Cells(Current_Row, 8).Value = Me.U7a_Vendor.Value
    
' [4] Update the Array for Combobox Display
    Dim PI_Array()
    ReDim PI_Array(n_proc, 4)
    
    For a = 1 To n_proc
        PI_Array(a - 1, 0) = Worksheets("B7").Cells(7 + n_feed + a, 2).Value
        PI_Array(a - 1, 1) = Worksheets("B7").Cells(7 + n_feed + a, 3).Value
        PI_Array(a - 1, 2) = Worksheets("B10").Cells(7 + n_feed + a, 4).Value
        PI_Array(a - 1, 3) = Format(Worksheets("O3").Cells(6 + a, 8).Value, "$0,000.00")
    Next a
    Me.U7a_IntervalSelect_Combobox.list = PI_Array
    
' [5] Check if Equipment Costs for all Intervals have been Specified, if so, then update the CHECKSUM
    Worksheets("O3").Range("F2").Value = 1
    For a = 1 To n_proc
        If Worksheets("O3").Cells(6 + a, 8).Value = "" Then
            Worksheets("O3").Range("F2").Value = 0
        End If
    Next a
    If Worksheets("O3").Range("F2").Value = 1 Then
        Worksheets("S7").Shapes.Range(Array("TextBox 21")).Select
        With Selection
            .Text = "¡î  -  Equipment Costs for all intervals have been specified."
            .Font.Italic = msoTrue
            .Font.Size = 12
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
        End With
    Else
        Worksheets("S7").Shapes.Range(Array("TextBox 21")).Select
        With Selection
            .Text = "X  -  Equipment Costs for all intervals have been specified."
            .Font.Italic = msoTrue
            .Font.Size = 12
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(118, 97, 66)
        .Transparency = 0
        .Solid
        End With
    End If
    Range("A1").Select

' [6] Equipment Cost Calculated!!
    MsgBox "Purchased Equipment Cost Saved!!", vbExclamation, "TIPEM- Notice"
End Sub





'**** CLOSE USERFORM *****
Private Sub U7a_Save_Click()
Unload Me
End
End Sub
