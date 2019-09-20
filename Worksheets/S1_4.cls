VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    On Error Resume Next
    Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = False
    Worksheets("S4").Shapes.Range(Array("Button 22")).Visible = False
    Worksheets("S4").Shapes.Range(Array("TextBox 69")).Visible = False
    
    '[1] ComboBox to Select Intervals
    Dim Num_Steps As Integer
    Num_Steps = Worksheets("S4").Cells(12, 8).Value + 2
    ' If Step = 1, AKA Feedstock Step, then:
    If S5_IntervalList.Column(0) = 1 Then
        S5_IntervalDisplay.Text = "Feedstock" & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    ' If Step = Last Step, AKA Product Step, then:
    ElseIf S5_IntervalList.Column(0) = Num_Steps Then
        S5_IntervalDisplay.Text = "Product" & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    ' Else, it is a process step
    Else
        S5_IntervalDisplay.Text = "Process Step " & S5_IntervalList.Column(0) & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    End If
    
    '[2] Assign Current Step/Interval to Workbook
    Worksheets("B10").Range("H3").Value = S5_IntervalList.Column(0)
    Worksheets("B10").Range("K3").Value = S5_IntervalList.Column(1)
    
    '[3] Display Current Interval Name in TextBox
    Dim row_source As Integer
    Num_Int = Sheets("S4").Range("H14").Value + 7
    
    For a = 8 To Num_Int
        If S5_IntervalList.Column(0) = Worksheets("B10").Cells(a, 2).Value And S5_IntervalList.Column(1) = Worksheets("B10").Cells(a, 3).Value Then
            row_source = a
        End If
    Next a
    
    If ActiveSheet.Name = "S4" Then
        Worksheets("S4").Shapes.Range(Array("TextBox 184")).Select
        If S5_IntervalList.Column(0) = 1 And Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "FEED INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
            '**FEEDSTOCK SPECIFICATION** If Feed Interval is Specified, then Hide Process Interval Figures
            Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = True
            Worksheets("S4").Shapes.Range(Array("Button 22")).Visible = True
        ElseIf S5_IntervalList.Column(0) = Num_Steps And Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PRODUCT INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
            '**FEEDSTOCK SPECIFICATION** If Feed Interval is Specified, then Hide Process Interval Figures
            Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = True
            Worksheets("S4").Shapes.Range(Array("TextBox 69")).Visible = True
        ElseIf Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PROCESS INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
        End If
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14).ParagraphFormat
            .Alignment = msoAlignCenter
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 9).Font
            .Bold = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(118, 97, 66)
            .Size = 16
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(10, 5).Font
            .Bold = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(118, 97, 66)
            .Size = 16
        End With
        Range("A1").Select
    End If
End Sub


Sub S5_IntervalList_Change()
    On Error Resume Next
    Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = False
    Worksheets("S4").Shapes.Range(Array("Button 22")).Visible = False
    Worksheets("S4").Shapes.Range(Array("TextBox 69")).Visible = False
    
    '[1] ComboBox to Select Intervals
    Dim Num_Steps As Integer
    Num_Steps = Worksheets("S4").Cells(12, 8).Value + 2
    ' If Step = 1, AKA Feedstock Step, then:
    If S5_IntervalList.Column(0) = 1 Then
        S5_IntervalDisplay.Text = "Feedstock" & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    ' If Step = Last Step, AKA Product Step, then:
    ElseIf S5_IntervalList.Column(0) = Num_Steps Then
        S5_IntervalDisplay.Text = "Product" & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    ' Else, it is a process step
    Else
        S5_IntervalDisplay.Text = "Process Step " & S5_IntervalList.Column(0) & "-" & S5_IntervalList.Column(1) & "   |   " & S5_IntervalList.Column(2)
    End If
    
    '[2] Assign Current Step/Interval to Workbook
    Worksheets("B10").Range("H3").Value = S5_IntervalList.Column(0)
    Worksheets("B10").Range("K3").Value = S5_IntervalList.Column(1)
    
    '[3] Display Current Interval Name in TextBox
    Dim row_source As Integer
    Num_Int = Sheets("S4").Range("H14").Value + 7
    
    For a = 8 To Num_Int
        If S5_IntervalList.Column(0) = Worksheets("B10").Cells(a, 2).Value And S5_IntervalList.Column(1) = Worksheets("B10").Cells(a, 3).Value Then
            row_source = a
        End If
    Next a
    
    If ActiveSheet.Name = "S4" Then
        Worksheets("S4").Shapes.Range(Array("TextBox 184")).Select
        If S5_IntervalList.Column(0) = 1 And Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "FEED INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
            '**FEEDSTOCK SPECIFICATION** If Feed Interval is Specified, then Hide Process Interval Figures
            Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = True
            Worksheets("S4").Shapes.Range(Array("Button 22")).Visible = True
        ElseIf S5_IntervalList.Column(0) = Num_Steps And Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PRODUCT INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
            '**FEEDSTOCK SPECIFICATION** If Feed Interval is Specified, then Hide Process Interval Figures
            Worksheets("S4").Shapes.Range(Array("Rectangle 3")).Visible = True
            Worksheets("S4").Shapes.Range(Array("TextBox 69")).Visible = True
        ElseIf Worksheets("B10").Cells(3, 5).Value = 1 Then
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PROCESS INTERVAL: " & Worksheets("B10").Cells(row_source, 4).Text
        End If
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14).ParagraphFormat
            .Alignment = msoAlignCenter
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 9).Font
            .Bold = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(118, 97, 66)
            .Size = 16
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(10, 5).Font
            .Bold = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(118, 97, 66)
            .Size = 16
        End With
        Range("A1").Select
    End If
End Sub


