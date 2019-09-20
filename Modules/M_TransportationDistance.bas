Attribute VB_Name = "M_TransportationDistance"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''     TIPEM TRANSPORTATIONG GEN   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ¿Ã¡ˆ»Ø/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [TRANSPORT] Generate Transportation Tables
Public Function TRANSPORT_Generate()
Application.ScreenUpdating = False
    '[1] Declare Variables
    Dim n_step As Integer
    Dim n_trans As Integer
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
    Dim cell_selected As Range
    Dim cell_selected_2 As Range
    Dim cell_selected_3 As Range
    
    n_step = Worksheets("S3").Range("H12").Value + 2
    n_trans = Worksheets("B5").Range("C1").Value
    n_total_interval = Worksheets("S3").Range("H14").Value
    n_feed_interval = Worksheets("S3").Range("F13").Value
    n_prod_interval = Worksheets("S3").Range("F" & 12 + n_step).Value
    n_proc_interval = n_total_interval - n_feed_interval - n_prod_interval
   
    '[2] Generate Tables
    For x = 1 To n_trans
       If x = 1 Then
          Set head = Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(1)
       Else
          Set head = Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(2)
       End If
       With head
          .Value = x & ") " & Worksheets("B5").Range("C" & 4 + x).Value
          .Font.Bold = True
       End With
               
       ' [Row Header for Process Step/Interval]
       Set cell_selected = Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(3)
       With cell_selected
          .Value = "Index"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       With cell_selected.Offset(1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       With cell_selected.Offset(, 1)
          .Value = "Step"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       With cell_selected.Offset(1, 1)
          .Value = "Interval"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       
       ' [Column Header for Process Step/Interval]
       Set cell_selected = Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(-2, 2)
       With Range(cell_selected, cell_selected.Offset(, n_total_interval - 1))
          .Merge
          .VerticalAlignment = xlCenter
          .HorizontalAlignment = xlCenter
          .Value = "Distance of Primary Streams (km)"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       Set cell_selected = Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(-2, 2 + n_total_interval)
       With Range(cell_selected, cell_selected.Offset(, n_total_interval - 1))
          .Merge
          .VerticalAlignment = xlCenter
          .HorizontalAlignment = xlCenter
          .Value = "Distance of Secondary (km)"
          .Font.Bold = True
          .Interior.Color = RGB(186, 244, 238)
       End With
       
       ' [Numbering Interval Index]
       For a = 1 To n_step
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          For b = 1 To n_interval
             With Worksheets("B11").Cells(5 + (x - 1) * (n_total_interval + 6), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = n_step_interval
               .Interior.Color = RGB(186, 244, 238)
             End With
             With Worksheets("B11").Cells(6 + (x - 1) * (n_total_interval + 6), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = b
               .Interior.Color = RGB(186, 244, 238)
             End With
          Next b
       Next a
       For a = 1 To n_step
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          For b = 1 To n_interval
             With Worksheets("B11").Cells(5 + (x - 1) * (n_total_interval + 6), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = n_step_interval
               .Interior.Color = RGB(186, 244, 238)
             End With
             With Worksheets("B11").Cells(6 + (x - 1) * (n_total_interval + 6), Columns.Count).End(xlToLeft).Offset(, 1)
               .Value = b
               .Interior.Color = RGB(186, 244, 238)
             End With
          Next b
       Next a
       
       For a = 1 To n_step
          n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
          n_interval = Worksheets("S3").Range("F" & 12 + a).Value
          For b = 1 To n_interval
             With Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).Offset(1)
               .Value = n_step_interval
               .Interior.Color = RGB(186, 244, 238)
             End With
             With Worksheets("B11").Cells(Rows.Count, "C").End(xlUp).Offset(1)
               .Value = b
               .Interior.Color = RGB(186, 244, 238)
             End With
          Next b
       Next a
       
       '[3] Drawing cell boundaries
       Worksheets("B11").Select
       Worksheets("B11").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
       Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
       Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
       Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
       Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
       Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
       Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
       Selection.HorizontalAlignment = xlCenter
       Selection.VerticalAlignment = xlCenter
    Next x
    Worksheets("S4").Select
Application.ScreenUpdating = True
End Function




' [TRANSPORT] Delete Transportation Tables
Public Function TRANSPORT_Delete()
Worksheets("B11").Cells.Clear
End Function


