Attribute VB_Name = "M_NetworkDrawing"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''     TIPEM PSIN DRAWING MODULE   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ¿Ã¡ˆ»Ø/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [INFRA] ReDefine System Size (Process Steps)
Public Sub S3_SpecifySize_1()
    ' *** PART 1*** Delete Current Network
    ' [1] Declare Variables
    Dim Shp As Shape
    Dim n_step As Integer
    Dim n_interval As Integer
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
    
    ' [3] Delete previous network figure
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp

    ' [4] Delete Process Interval Specification Table for Step 5
    Application.ScreenUpdating = False
    TIPEM_Delete_IntervalSpecTable
    TRANSPORT_Delete
    Worksheets("S3").Activate
    Application.ScreenUpdating = True
    
    ' *** PART 2*** Define System Size
    ' [1] Define Step Number and Range
    Dim number_step As Integer
    Dim cell_selected As Range
    
    ' [2] Find end of Range
    number_step = Range("H12").Value
    Set cell_selected = Cells(15 + number_step, 6)
    
    ' [3] Select Beginning of Range and Clear Contents
    Range(Cells(13, 4), cell_selected).ClearContents
    Range(Cells(13, 4), cell_selected).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

    ' [4] Clear Any Border Styles of Range
    Range(Cells(13, 4), cell_selected).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    ' [5] Clear Step Count
    Range("H12").ClearContents
    Range("A1").Select
    
    ' [6] Show Userform
    U4a_System_Size.Show
End Sub





' [PROCESS DRAWING] Define Stream Connections **FRESH START**
Sub S3_DefineConnections()
    On Error Resume Next
    ' Make sure intervals have been defined
    For aa = 1 To 22
        If Worksheets("S3").Cells(12 + aa, 6).Value = "Enter Interval #" Then
            MsgBox "Make sure all process intervals have been specified!!", vbExclamation, "TIPEM- Error"
            Exit Sub
        End If
    Next aa

    ' If Connectivity Matrix Already has Values (aka, already has User-defined Connections), then dont re-draw the matrix
    If WorksheetFunction.CountA(Worksheets("B7").Range("D8:CZ220")) <> 0 Then
        U4b_Stream_Connections.Show vbModeless
        Exit Sub
    End If

    ' *** FIRST PART IS TO DELETE ANY EXISTING CONNECTIVITY MATRIXES***
    ' [1] Define Variables
    Application.ScreenUpdating = False
    Application.Goto (Sheets("B7").Range("B4:CZ220"))
    Worksheets("B7").Range("B4:CZ220").ClearContents
    Application.Goto (Sheets("B12").Range("B4:CZ220"))
    Worksheets("B12").Range("B4:CZ220").ClearContents
    Selection.Font.Bold = False
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Interior.TintAndShade = 0
    
    ' *** SECOND PART IS TO GENERATE CONNECTIVITY MATRIXES***
    ' [1] Define Variables
    Dim n_step As Integer
    Dim n_total_interval As Integer
    Dim n_feed_interval As Integer
    Dim n_proc_interval As Integer
    Dim n_prod_interval As Integer
    Dim n_step_interval As Integer
    Dim n_interval As Integer
    Dim a As Integer
    Dim b As Integer
    Dim n As Integer
    Dim x As Integer
    Dim head As Range
    Dim cell_selected_1 As Range
    Dim cell_selected_2 As Range
    Dim cell_selected_3 As Range
    
    ' [2] Define Ranges
    n_step = Worksheets("S3").Range("H12").Value + 2
    n_total_interval = Worksheets("S3").Range("H14").Value
    n_feed_interval = Worksheets("S3").Range("F13").Value
    n_prod_interval = Worksheets("S3").Range("F" & 12 + n_step).Value
    n_proc_interval = n_total_interval - n_feed_interval - n_prod_interval
    
    ' [3] If no Matrix Connectivity Table exists, then
    For x = 1 To 2
       ' [3a] Generate Title
       If x = 1 Then
          Worksheets("B7").Activate
          Set head = Cells(Rows.Count, "B").End(xlUp).Offset(1)
       Else
          Set head = Cells(Rows.Count, "B").End(xlUp).Offset(2)
       End If
       If x = 1 Then
          With head
             .Value = "PRIMARY PROCESS STREAMS"
             .Font.Bold = True
             .HorizontalAlignment = xlLeft
          End With
       Else
          With head
             .Value = "SECONDARY PROCESS STREAMS"
             .Font.Bold = True
             .HorizontalAlignment = xlLeft
          End With
       End If
       
       ' [3b] Generate Matrix Table Headers
       Worksheets("B7").Activate
       Set cell_selected_1 = Cells(Rows.Count, "B").End(xlUp).Offset(2)
       With cell_selected_1
          .Value = "Index"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(, 1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(1, 1)
          .Value = "Interval"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       
       ' [3c] Generate Interval and Step Dimensions
       For a = 1 To n_step
          ' Assign Step Name to Variable
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          ' Assign # of Intervals per Step to Variable
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          ' Generate Matrix Table Headers at Row 6 & 7 of Sheet B7
          For b = 1 To n_interval
             With Worksheets("B7").Cells(6 + (x - 1) * (n_total_interval + 5), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = n_step_interval
               .Interior.Color = RGB(221, 235, 247)
             End With
             With Worksheets("B7").Cells(7 + (x - 1) * (n_total_interval + 5), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = b
               .Interior.Color = RGB(221, 235, 247)
             End With
          Next b
       Next a
       
       ' [3d] Generate Matrix
       For a = 1 To n_step
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          For b = 1 To n_interval
             Worksheets("B7").Activate
             With Cells(Rows.Count, "B").End(xlUp).Offset(1)
               .Value = n_step_interval
               .Interior.Color = RGB(221, 235, 247)
             End With
             With Cells(Rows.Count, "C").End(xlUp).Offset(1)
               .Value = b
               .Interior.Color = RGB(221, 235, 247)
             End With
          Next b
       Next a
       
       ' [3e] Draw Cell Boundaries
       Worksheets("B7").Activate
       Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
       Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
       Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
       Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
       Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
       Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
       Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
       Selection.HorizontalAlignment = xlCenter
       Selection.VerticalAlignment = xlCenter
       
       ' [3f] Coloring inactive cells
       For a = 1 To n_total_interval
          ' Assign current matrix to selected sell
          Set cell_selected_1 = Worksheets("B7").Cells(7 + a, 4)
          Set cell_selected_2 = Worksheets("B7").Cells(7 + a, 3 + a)
          Range(cell_selected_1, cell_selected_2).Select
          Selection.Interior.Color = 255
          
          ' Assign new matrix for secondary streams to selected cell
          Set cell_selected_1 = Worksheets("B7").Cells(n_total_interval + 12 + a, 4)
          Set cell_selected_2 = Worksheets("B7").Cells(n_total_interval + 12 + a, 3 + a)
          Range(cell_selected_1, cell_selected_2).Select
          Selection.Interior.Color = 255
       Next a
    Next x
    
    ' [4] Copy Matrix to Sheet B12 for Pathway Specification
    For x = 1 To 2
       ' [4a] Generate Title
       If x = 1 Then
          Worksheets("B12").Activate
          Set head = Cells(Rows.Count, "B").End(xlUp).Offset(1)
       Else
          Set head = Cells(Rows.Count, "B").End(xlUp).Offset(2)
       End If
       If x = 1 Then
          With head
             .Value = "PRIMARY PROCESS STREAMS"
             .Font.Bold = True
             .HorizontalAlignment = xlLeft
          End With
       Else
          With head
             .Value = "SECONDARY PROCESS STREAMS"
             .Font.Bold = True
             .HorizontalAlignment = xlLeft
          End With
       End If
       
       ' [3b] Generate Matrix Table Headers
       Worksheets("B12").Activate
       Set cell_selected_1 = Cells(Rows.Count, "B").End(xlUp).Offset(2)
       With cell_selected_1
          .Value = "Index"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(, 1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       With cell_selected_1.Offset(1, 1)
          .Value = "Interval"
          .Font.Bold = True
          .Interior.Color = RGB(221, 235, 247)
       End With
       
       ' [3c] Generate Interval and Step Dimensions
       For a = 1 To n_step
          ' Assign Step Name to Variable
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          ' Assign # of Intervals per Step to Variable
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          ' Generate Matrix Table Headers at Row 6 & 7 of Sheet B7
          For b = 1 To n_interval
             With Worksheets("B12").Cells(6 + (x - 1) * (n_total_interval + 5), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = n_step_interval
               .Interior.Color = RGB(221, 235, 247)
             End With
             With Worksheets("B12").Cells(7 + (x - 1) * (n_total_interval + 5), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = b
               .Interior.Color = RGB(221, 235, 247)
             End With
          Next b
       Next a
       
       ' [3d] Generate Matrix
       For a = 1 To n_step
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          For b = 1 To n_interval
             Worksheets("B12").Activate
             With Cells(Rows.Count, "B").End(xlUp).Offset(1)
               .Value = n_step_interval
               .Interior.Color = RGB(221, 235, 247)
             End With
             With Cells(Rows.Count, "C").End(xlUp).Offset(1)
               .Value = b
               .Interior.Color = RGB(221, 235, 247)
             End With
          Next b
       Next a
       
       ' [3e] Draw Cell Boundaries
       Worksheets("B12").Activate
       Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
       Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
       Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
       Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
       Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
       Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
       Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
       Selection.HorizontalAlignment = xlCenter
       Selection.VerticalAlignment = xlCenter
       
       ' [3f] Coloring inactive cells
       For a = 1 To n_total_interval
          ' Assign current matrix to selected sell
          Set cell_selected_1 = Worksheets("B12").Cells(7 + a, 4)
          Set cell_selected_2 = Worksheets("B12").Cells(7 + a, 3 + a)
          Range(cell_selected_1, cell_selected_2).Select
          Selection.Interior.Color = 255
          
          ' Assign new matrix for secondary streams to selected cell
          Set cell_selected_1 = Worksheets("B12").Cells(n_total_interval + 12 + a, 4)
          Set cell_selected_2 = Worksheets("B12").Cells(n_total_interval + 12 + a, 3 + a)
          Range(cell_selected_1, cell_selected_2).Select
          Selection.Interior.Color = 255
       Next a
    Next x
    
    Worksheets("S3").Activate
    Application.ScreenUpdating = True
    U4b_Stream_Connections.Show vbModeless
End Sub


' [PROCESS DRAWING] Draw Network
Sub S3_DrawSystem()
    Application.EnableEvents = False
    On Error Resume Next
    ' Make sure intervals have been defined
    For aa = 1 To 22
        If Worksheets("S3").Cells(12 + aa, 6).Value = "Enter Interval #" Then
            MsgBox "Make sure all process intervals have been specified!!", vbExclamation, "TIPEM- Error"
            Exit Sub
        End If
    Next aa
    ' **NOTE** This will delete all current stream connection and connectivity information!!
    Warning1 = MsgBox("This will delete ALL current Stream Connection and Connectivity Information as well as ALL related Process Specification information!!!", vbOKCancel, "TIPEM- Warning")
    If Warning1 = vbCancel Then
        Exit Sub
    End If
    
    ' [1] Delete all Connectivity Information
    Dim n_interval As Integer
    Dim n_mat As Integer
    n_interval = Worksheets("S3").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value
    
    Worksheets("B9").Range("B4:F2000").ClearContents
    Worksheets("B8").Range("B4:F2000").ClearContents
    Worksheets("B7").Range("D8:CZ220").ClearContents
    Worksheets("B12").Range("D8:CZ220").ClearContents
    Worksheets("B11").Cells.Clear

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
    
    ' Mass Balance Available Checksum
    Worksheets("O1").Range("F2").Value = 0
    Worksheets("O2").Range("F2").Value = 0
    Worksheets("O3").Range("F2").Value = 0
    Worksheets("O4").Range("F2").Value = 0
    Worksheets("O4").Range("H2").Value = 0
    
    
    ' ** DRAW SYSTEM CODE STARTS HERE **
    ' [1] Define Variables
    Dim Shp As Shape
    Dim n_step As Integer
    Dim Current_Step As Integer
    Dim current_interval As Integer
    Dim max_interval As Integer
    Dim connection_p() As Integer
    Dim connection_s() As Integer
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim x As Integer
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S3").Range("H12").Value + 2
    
    ' [3] Change Variable Dimension to n_int by n_int Array
    ReDim connection_p(1 To n_interval, 1 To n_interval)
    ReDim connection_s(1 To n_interval, 1 To n_interval)
    
    ' [4] Delete any previous Networks
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
       
    ' [5] Recall Connectivity Information
    For a = 1 To n_interval
       For b = 1 To n_interval
          connection_p(a, b) = Worksheets("B7").Cells(7 + a, 3 + b).Value
          connection_s(a, b) = Worksheets("B7").Cells(12 + n_interval + a, 3 + b).Value
       Next b
    Next a
    
    ' [6] Find maximum processing interval number
    max_interval = 0
    For a = 1 To n_step
       x = Worksheets("S3").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [7] Network Drawing Module [MAIN]
    x = 1
    For a = 1 To n_step
        ' Index steps and # of intervals in each step
        Current_Step = Worksheets("S3").Range("E" & 12 + a).Value
        current_interval = Worksheets("S3").Range("F" & 12 + a).Value
        RM_interval = Worksheets("S3").Range("F" & 13).Value
        
        ' For each interval:
        For b = 1 To current_interval
            'Draw shapes
            ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 180, 30, 75, 50).Select
            'Set shape style
            Selection.Name = "shape" & x
            With Selection.ShapeRange.Fill
              .Visible = msoTrue
              .ForeColor.RGB = RGB(217, 225, 213)
              .ForeColor.TintAndShade = 0
              .ForeColor.Brightness = 0
              .Transparency = 0
              .Solid
            End With
            With Selection.ShapeRange.Line
              .Visible = msoTrue
              .ForeColor.RGB = RGB(118, 97, 66)
              .ForeColor.TintAndShade = 0
              .ForeColor.Brightness = 0
              .Transparency = 0
              .Weight = 2
            End With
            
                '***Color Raw Material and Product Intervals Separately***
                ' Raw Material Intervals
                If a = 1 Then Selection.ShapeRange.Fill.ForeColor.RGB = RGB(228, 146, 118)
                If a = 1 Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(50, 50, 50)
                ' Product Intervals
                If a = n_step Then Selection.ShapeRange.Fill.ForeColor.RGB = RGB(133, 161, 209)
                If a = n_step Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(50, 50, 50)
            
            'Insert text in shapes
            Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
            
            'Name the shape according to index
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = a & "-" & b
                '***Name Raw Material and Product Intervals Separately***
                ' Raw Material Intervals
                If a = 1 Then Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "FEED" & "-" & b
                ' Product Intervals
                If a = n_step Then Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PROD" & "-" & b
                
            With Selection.ShapeRange(1).TextFrame2.TextRange.ParagraphFormat
              .FirstLineIndent = 0
              .Alignment = msoAlignCenter
            End With
            With Selection.ShapeRange(1).TextFrame2.TextRange.Font
              .NameComplexScript = "+mn-cs"
              .NameFarEast = "+mn-ea"
              .Fill.Visible = msoTrue
              .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
              .Fill.ForeColor.TintAndShade = 0
              .Fill.ForeColor.Brightness = 0
              .Fill.Transparency = 0
              .Fill.Solid
              .Size = 14
              .Name = "+mn-lt"
              .Bold = msoTrue
            End With
            
            'Change Font Size According to Syze of System
            If Worksheets("S3").Range("H12").Value >= 7 Then
               Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
            End If
            If Worksheets("S3").Range("H12").Value >= 10 Then
               Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
            End If
            
            'Locate shapes
            ' *** IF MAX INTERVAL IS 1 ***
            If max_interval = 1 Then
            Set Shp = ActiveSheet.Shapes("shape" & x)
            ' [NEW] Auto-Adjust Height between each Interval
            ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
            IntHeight = 480 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
            IntWidth = 1230 / ((2 * n_step) + (1.5 * (n_step - 1)))
            With Shp
            ' [NEW] Auto-Adjust the Shape of each process block according to the size of the system
               .Height = 250 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
               .Width = 1640 / ((2 * n_step) + (1.5 * (n_step - 1)))
            ' [NEW] Set the Position of Each Process Block
               .Top = 225 + 2 * IntHeight * (b - 1)
               .Left = 390 + 2 * IntWidth * (Current_Step - 1)
            End With
            ' *** IF MAX INTERVAL IS NOT 1 ***
            Else
                Set Shp = ActiveSheet.Shapes("shape" & x)
                ' [NEW] Auto-Adjust Height between each Interval
                ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
                IntHeight = 450 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
                IntWidth = 1200 / ((2 * n_step) + (1.5 * (n_step - 1)))
                With Shp
                ' [NEW] Auto-Adjust the Shape of each process block according to the size of the system
                   .Height = 640 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
                   .Width = 1640 / ((2 * n_step) + (1.5 * (n_step - 1)))
                ' [NEW] Set the Position of Each Process Block
                   .Top = 225 + 2 * IntHeight * (b - 1)
                   .Left = 390 + 2 * IntWidth * (Current_Step - 1)
                End With
            End If
            ' Iterate for Next Loop
            x = x + 1
        Next b
    Next a
      
    ' [8] Draw arrows to connect intevals
    ' Primary connection
    For a = 1 To n_interval
       For b = 1 To n_interval
          If connection_p(a, b) = 1 Then
             ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 112.5, 225, 180, 225).Select
             With Selection.ShapeRange
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                .ConnectorFormat.BeginConnect ActiveSheet.Shapes("shape" & a), 4
                .ConnectorFormat.EndConnect ActiveSheet.Shapes("shape" & b), 2
                .ShapeStyle = msoLineStylePreset8
                .Line.ForeColor.RGB = RGB(0, 112, 192)
                .Line.Visible = msoTrue
                .Line.Transparency = 0
                .Line.Weight = 1.2
                .Name = "Pri_arrow " & a & "-" & b
             End With
          End If
       Next b
    Next a

    ' Secondary connection
    For a = 1 To n_interval
       For b = 1 To n_interval
          If connection_s(a, b) = 1 Then
             ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 112.5, 225, 180, 225).Select
             With Selection.ShapeRange
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                .ConnectorFormat.BeginConnect ActiveSheet.Shapes("shape" & a), 4
                .ConnectorFormat.EndConnect ActiveSheet.Shapes("shape" & b), 2
                .ShapeStyle = msoLineStylePreset8
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                .Line.Visible = msoTrue
                .Line.Transparency = 0
                .Line.Weight = 1.2
                .Name = "Pri_arrow " & a & "-" & b
             End With
          End If
       Next b
    Next a
    
    ' [9] Initialize Process Interval Specification Table for Step 5
    Application.ScreenUpdating = False
    TIPEM_Delete_IntervalSpecTable
    TIPEM_Create_IntervalSpecTable
    
    ' [10] Delete and Regenerate Transportation Matrixes
    TRANSPORT_Generate
    Worksheets("S3").Activate
    Range("A1").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub





' [PROCESS DRAWING] **PRIMARY** Find the Row of the Stream Connectivity Matching Input from ComboBox in U4b_Stream_Connections
Function Network_Drawing_Pathway_Config()
    ' [1] Define Variables
    Dim Shp As Shape
    Dim n_step As Integer
    Dim n_interval As Integer
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim max_interval As Integer
    Dim connection_p() As Integer
    Dim connection_s() As Integer
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim x As Integer
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S3").Range("H12").Value + 2
    n_interval = Worksheets("S3").Range("H14").Value
    
    ' [3] Change Variable Dimension to n_int by n_int Array
    ReDim connection_p(1 To n_interval, 1 To n_interval)
    ReDim connection_s(1 To n_interval, 1 To n_interval)
    
    ' [4] Delete any previous Networks
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
       
    ' [5] Recall Connectivity Information
    For a = 1 To n_interval
       For b = 1 To n_interval
          connection_p(a, b) = Worksheets("B12").Cells(7 + a, 3 + b).Value
          connection_s(a, b) = Worksheets("B12").Cells(12 + n_interval + a, 3 + b).Value
       Next b
    Next a
    
    ' [6] Find maximum processing interval number
    max_interval = 0
    For a = 1 To n_step
       x = Worksheets("S3").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [7] Network Drawing Module
    x = 1
    For a = 1 To n_step
       ' Index steps and # of intervals in each step
       Current_Step = Worksheets("S3").Range("E" & 12 + a).Value
       Current_Int = Worksheets("S3").Range("F" & 12 + a).Value
       ' For each interval:
       For b = 1 To Current_Int
          'Draw shapes
          Worksheets("S8").Shapes.AddShape(msoShapeRoundedRectangle, 180, 30, 75, 50).Select
          'Set shape style
          Selection.Name = "shape" & x
          With Selection.ShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(217, 225, 213)
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
          End With
          With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(118, 97, 66)
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Weight = 2
          End With
                '***Color Raw Material and Product Intervals Separately***
                ' Raw Material Intervals
                If a = 1 Then Selection.ShapeRange.Fill.ForeColor.RGB = RGB(228, 146, 118)
                If a = 1 Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(50, 50, 50)
                ' Product Intervals
                If a = n_step Then Selection.ShapeRange.Fill.ForeColor.RGB = RGB(133, 161, 209)
                If a = n_step Then Selection.ShapeRange.Line.ForeColor.RGB = RGB(50, 50, 50)
          'Insert text in shapes
          Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
          
          'Name the shape according to index
          Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = a & "-" & b
                '***Name Raw Material and Product Intervals Separately***
                ' Raw Material Intervals
                If a = 1 Then Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "FEED" & "-" & b
                ' Product Intervals
                If a = n_step Then Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PROD" & "-" & b
          With Selection.ShapeRange(1).TextFrame2.TextRange.ParagraphFormat
            .FirstLineIndent = 0
            .Alignment = msoAlignCenter
          End With
          With Selection.ShapeRange(1).TextFrame2.TextRange.Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
            .Fill.ForeColor.TintAndShade = 0
            .Fill.ForeColor.Brightness = 0
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Name = "+mn-lt"
            .Bold = msoTrue
          End With
       
          'Change Font Size According to Syze of System
          If Worksheets("S3").Range("H12").Value >= 7 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
          End If
          If Worksheets("S3").Range("H12").Value >= 10 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
          End If
       
          'Locate shapes
          ' *** IF MAX INTERVAL IS 1 ***
          If max_interval = 1 Then
          Set Shp = Worksheets("S8").Shapes("shape" & x)
          ' [NEW] Auto-Adjust Height between each Interval
                ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
                IntHeight = 520 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
                IntWidth = 1400 / ((2 * n_step) + (1.5 * (n_step - 1)))
          With Shp
          ' [NEW] Auto-Adjust the Shape of each process block according to the size of the system
                .Height = 600 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
                .Width = 1700 / ((2 * n_step) + (1.5 * (n_step - 1)))
          ' [NEW] Set the Position of Each Process Block
                .Top = 200 - 2 * IntHeight * (b - 1)
                .Left = 300 + 2 * IntWidth * (Current_Step - 1)
          End With
          ' *** IF MAX INTERVAL IS NOT 1 ***
          Else
              Set Shp = Worksheets("S8").Shapes("shape" & x)
              ' [NEW] Auto-Adjust Height between each Interval
              ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
              IntHeight = 520 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
              IntWidth = 1400 / ((2 * n_step) + (1.5 * (n_step - 1)))
              With Shp
              ' [NEW] Auto-Adjust the Shape of each process block according to the size of the system
                 .Height = 600 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
                 .Width = 1700 / ((2 * n_step) + (1.5 * (n_step - 1)))
              ' [NEW] Set the Position of Each Process Block
                 .Top = 200 + 2 * IntHeight * (b - 1)
                 .Left = 300 + 2 * IntWidth * (Current_Step - 1)
              End With
          End If
          ' Iterate for Next Loop
          x = x + 1
       Next b
    Next a
      
    'Draw arrows to connect intevals
    'Primary connection
    For a = 1 To n_interval
       For b = 1 To n_interval
          If connection_p(a, b) = 1 Then
             Worksheets("S8").Shapes.AddConnector(msoConnectorStraight, 112.5, 225, 180, 225).Select
             With Selection.ShapeRange
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                .ConnectorFormat.BeginConnect Worksheets("S8").Shapes("shape" & a), 4
                .ConnectorFormat.EndConnect Worksheets("S8").Shapes("shape" & b), 2
                .ShapeStyle = msoLineStylePreset8
                .Line.ForeColor.RGB = RGB(0, 112, 192)
                .Line.Visible = msoTrue
                .Line.Transparency = 0
                .Line.Weight = 1.2
                .Name = "Pri_arrow " & a & "-" & b
             End With
          End If
       Next b
    Next a

    'Secondary connection
    For a = 1 To n_interval
       For b = 1 To n_interval
          If connection_s(a, b) = 1 Then
             Worksheets("S8").Shapes.AddConnector(msoConnectorStraight, 112.5, 225, 180, 225).Select
             With Selection.ShapeRange
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                .ConnectorFormat.BeginConnect Worksheets("S8").Shapes("shape" & a), 4
                .ConnectorFormat.EndConnect Worksheets("S8").Shapes("shape" & b), 2
                .ShapeStyle = msoLineStylePreset8
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                .Line.Visible = msoTrue
                .Line.Transparency = 0
                .Line.Weight = 1.2
                .Name = "Pri_arrow " & a & "-" & b
             End With
          End If
       Next b
    Next a
End Function





' [PROCESS DRAWING] **PRIMARY** Find the Row of the Stream Connectivity Matching Input from ComboBox in U4b_Stream_Connections
Function ConFind(SStep As Long, SInt As Long, DStep As Long, DInt As Long) As Long
    ' Declare Variables
    Dim Row As Long
    For Row = 4 To 5000
      FindSStep = Val(Sheets("B8").Cells(Row, 3).Value)
      FindSInt = Val(Sheets("B8").Cells(Row, 4).Value)
      FindDStep = Val(Sheets("B8").Cells(Row, 5).Value)
      FindDInt = Val(Sheets("B8").Cells(Row, 6).Value)
      If FindSStep = SStep And FindSInt = SInt And FindDStep = DStep And FindDInt = DInt Then
        ConFind = Row
        Exit Function
      End If
    Next Row
End Function
' [PROCESS DRAWING] **SECONDARY** Find the Row of the Stream Connectivity Matching Input from ComboBox in U4b_Stream_Connections
Function ConFind2(SStep As Long, SInt As Long, DStep As Long, DInt As Long) As Long
    ' Declare Variables
    Dim Row As Long
    For Row = 4 To 5000
      FindSStep = Val(Sheets("B9").Cells(Row, 3).Value)
      FindSInt = Val(Sheets("B9").Cells(Row, 4).Value)
      FindDStep = Val(Sheets("B9").Cells(Row, 5).Value)
      FindDInt = Val(Sheets("B9").Cells(Row, 6).Value)
      If FindSStep = SStep And FindSInt = SInt And FindDStep = DStep And FindDInt = DInt Then
        ConFind2 = Row
        Exit Function
      End If
    Next Row
End Function



