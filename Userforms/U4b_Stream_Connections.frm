VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U4b_Stream_Connections 
   Caption         =   "Define Stream Connections"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11850
   OleObjectBlob   =   "U4b_Stream_Connections.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U4b_Stream_Connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'************ DYNAMIC COMBOBOX FOR FINDING SOURCE/DESTINATION STEP/INT *********
'*******************************************************************************
' Create Private Dictionaries for SourceStep-SourceInterval and DestStep-DestInterval
Private SourceDic As Object
Private DestDic As Object
Private SourceDic_S As Object
Private DestDic_S As Object



Private Sub MultiPage1_Change()

End Sub

' Display Via Textbox, what Source/Dest Intervals are Selected
Private Sub U4b_ConnectionsList_Change()
    If Val(U4b_ConnectionsList.Column(0)) = 1 Then
        U4b_DisplayPrimaryConnections.Text = "Source FEED" & "-" & U4b_ConnectionsList.Column(1) & "    to    " & "Destination " & U4b_ConnectionsList.Column(2) & "-" & U4b_ConnectionsList.Column(3)
    ElseIf Val(U4b_ConnectionsList.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U4b_DisplayPrimaryConnections.Text = "Source " & U4b_ConnectionsList.Column(0) & "-" & U4b_ConnectionsList.Column(1) & "    to    " & "Destination PROD" & "-" & U4b_ConnectionsList.Column(3)
    Else
    U4b_DisplayPrimaryConnections.Text = "Source " & U4b_ConnectionsList.Column(0) & "-" & U4b_ConnectionsList.Column(1) & "    to    " & "Destination " & U4b_ConnectionsList.Column(2) & "-" & U4b_ConnectionsList.Column(3)
    End If
End Sub

Private Sub U4b_ConnectionsList_S_Change()
    If Val(U4b_ConnectionsList_S.Column(0)) = 1 Then
        U4b_DisplaySecondaryConnections.Text = "Source FEED" & "-" & U4b_ConnectionsList_S.Column(1) & "    to    " & "Destination " & U4b_ConnectionsList_S.Column(2) & "-" & U4b_ConnectionsList_S.Column(3)
    ElseIf Val(U4b_ConnectionsList_S.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U4b_DisplaySecondaryConnections.Text = "Source " & U4b_ConnectionsList_S.Column(0) & "-" & U4b_ConnectionsList_S.Column(1) & "    to    " & "Destination PROD" & "-" & U4b_ConnectionsList_S.Column(3)
    Else
    U4b_DisplaySecondaryConnections.Text = "Source " & U4b_ConnectionsList_S.Column(0) & "-" & U4b_ConnectionsList_S.Column(1) & "    to    " & "Destination " & U4b_ConnectionsList_S.Column(2) & "-" & U4b_ConnectionsList_S.Column(3)
    End If
End Sub



' Initialize Userform. Launch code to load Interval Values based on Source Values
Private Sub UserForm_Initialize()
    S4_SourceList
    S4_SourceList_S
    S4_DestList
    S4_DestList_S
End Sub



' Procedure for Populating only Unique Process Steps for Source Step
Sub S4_SourceList()
    ' [1] Declare Variables
    Dim SourceR As Range
    Dim list As Object
    
    ' [2] Populate Dictionary
    Set SourceDic = CreateObject("Scripting.Dictionary")
    
    ' [3] The dictionary uses the Secondary Connection Matrix to populate since it is below the primary matrix and xlUp is used. Thus we need to find the index of the 1st row of the Secondary Connection Matrix by referencing total number of intervals + offset
    total_intervals = Sheets("S4").Range("H14").Value + 7

    ' [3] Assign only Unique Values to Dictionary for Source Step ComboBox
    With Worksheets("B7")
        For Each SourceR In .Range("B8", .Cells(total_intervals, "B"))
            If Not SourceDic.exists(SourceR.Text) Then
                Set list = CreateObject("System.Collections.ArrayList")
                SourceDic.Add SourceR.Text, list
            End If
            If Not SourceDic(SourceR.Text).Contains(SourceR.Offset(0, 1).Value) Then
                SourceDic(SourceR.Text).Add SourceR.Offset(0, 1).Value
            End If
        Next
    End With

    ' [4] Load Assigned Values into ComboBox
    U4b_New_Combo_SourceStep.list = SourceDic.keys
    U4b_New_Combo_SourceInt.Clear
End Sub

Sub S4_SourceList_S()
    ' [1] Declare Variables
    Dim SourceR_S As Range
    Dim list As Object
    
    ' [2] Populate Dictionary
    Set SourceDic_S = CreateObject("Scripting.Dictionary")
    
    ' [3] The dictionary uses the Secondary Connection Matrix to populate since it is below the primary matrix and xlUp is used. Thus we need to find the index of the 1st row of the Secondary Connection Matrix by referencing total number of intervals + offset
    total_intervals_S = Sheets("S4").Range("H14").Value + 13

    ' [4] Assign only Unique Values to Dictionary for Source Step ComboBox
    With Worksheets("B7")
        For Each SourceR_S In .Range(.Cells(total_intervals_S, "B"), .Cells((2 * total_intervals_S) - 13, "B"))
            If Not SourceDic_S.exists(SourceR_S.Text) Then
                Set list = CreateObject("System.Collections.ArrayList")
                SourceDic_S.Add SourceR_S.Text, list
            End If
            If Not SourceDic_S(SourceR_S.Text).Contains(SourceR_S.Offset(0, 1).Value) Then
                SourceDic_S(SourceR_S.Text).Add SourceR_S.Offset(0, 1).Value
            End If
        Next
    End With

    ' [5] Load Assigned Values into ComboBox
    U4b_Sec_Combo_SourceStep.list = SourceDic_S.keys
    U4b_Sec_Combo_SourceInt.Clear
End Sub



' Procedure for Populating only Unique Process Steps for Dest Step
Sub S4_DestList()
    ' [1] Declare Variables
    Dim DestR As Range
    Dim list As Object
    
    ' [2] Populate Dictionary
    Set DestDic = CreateObject("Scripting.Dictionary")

    ' [3] Assign only Unique Values to Dictionary for Source Step ComboBox
    With Worksheets("B7")
        For Each DestR In .Range("D6", .Cells(6, .Columns.Count).End(xlToLeft))
            If Not DestDic.exists(DestR.Text) Then
                Set list = CreateObject("System.Collections.ArrayList")
                DestDic.Add DestR.Text, list
            End If
            If Not DestDic(DestR.Text).Contains(DestR.Offset(1, 0).Value) Then
                DestDic(DestR.Text).Add DestR.Offset(1, 0).Value
            End If
        Next
    End With

    ' [4] Load Assigned Values into ComboBox
    U4b_New_Combo_DestStep.list = DestDic.keys
    U4b_New_Combo_DestInt.Clear
End Sub

Sub S4_DestList_S()
    ' [1] Declare Variables
    Dim DestR_S As Range
    Dim list As Object
    
    ' [2] Populate Dictionary
    Set DestDic_S = CreateObject("Scripting.Dictionary")

    ' [3] The dictionary uses the Secondary Connection Matrix to populate since it is below the primary matrix and xlUp is used. Thus we need to find the index of the 1st row of the Secondary Connection Matrix by referencing total number of intervals + offset
    total_intervals_S = Sheets("S4").Range("H14").Value + 13
    
    ' [4] Assign only Unique Values to Dictionary for Source Step ComboBox
    With Worksheets("B7")
        For Each DestR_S In .Range(.Cells(total_intervals_S - 2, "D"), .Cells(total_intervals_S - 2, .Columns.Count).End(xlToLeft))
            If Not DestDic_S.exists(DestR_S.Text) Then
                Set list = CreateObject("System.Collections.ArrayList")
                DestDic_S.Add DestR_S.Text, list
            End If
            If Not DestDic_S(DestR_S.Text).Contains(DestR_S.Offset(1, 0).Value) Then
                DestDic_S(DestR_S.Text).Add DestR_S.Offset(1, 0).Value
            End If
        Next
    End With

    ' [5] Load Assigned Values into ComboBox
    U4b_Sec_Combo_DestStep.list = DestDic_S.keys
    U4b_Sec_Combo_DestInt.Clear
End Sub



' Procedure for Populating Source Interval based on Source Step Value
Private Sub U4b_New_Combo_SourceStep_Change()
    If U4b_New_Combo_SourceStep.ListIndex > -1 Then
        U4b_New_Combo_SourceInt.Clear
        SourceDic(U4b_New_Combo_SourceStep.Text).Sort
        U4b_New_Combo_SourceInt.list = SourceDic(U4b_New_Combo_SourceStep.Text).ToArray
    End If
End Sub

Private Sub U4b_Sec_Combo_SourceStep_Change()
    If U4b_Sec_Combo_SourceStep.ListIndex > -1 Then
        U4b_Sec_Combo_SourceInt.Clear
        SourceDic_S(U4b_Sec_Combo_SourceStep.Text).Sort
        U4b_Sec_Combo_SourceInt.list = SourceDic_S(U4b_Sec_Combo_SourceStep.Text).ToArray
    End If
End Sub



' Procedure for Populating Source Interval based on Destination Step Value
Private Sub U4b_New_Combo_DestStep_Change()
    If U4b_New_Combo_DestStep.ListIndex > -1 Then
        U4b_New_Combo_DestInt.Clear
        DestDic(U4b_New_Combo_DestStep.Text).Sort
        U4b_New_Combo_DestInt.list = DestDic(U4b_New_Combo_DestStep.Text).ToArray
    End If
End Sub

Private Sub U4b_Sec_Combo_DestStep_Change()
    If U4b_Sec_Combo_DestStep.ListIndex > -1 Then
        U4b_Sec_Combo_DestInt.Clear
        DestDic_S(U4b_Sec_Combo_DestStep.Text).Sort
        U4b_Sec_Combo_DestInt.list = DestDic_S(U4b_Sec_Combo_DestStep.Text).ToArray
    End If
End Sub



' Upon clicking "Add Stream Connection", update connectivity matrix, then update network drawing
Private Sub U4b_Button_Add_Click()
    ' *** PART1 *** Check if all ComboBox Fields are Specified
    ' [1] Check that all Input Entries are Identified
    If Val(U4b_New_Combo_SourceStep.Value) = 0 Or Val(U4b_New_Combo_SourceInt.Value) = 0 Or Val(U4b_New_Combo_DestStep.Value) = 0 Or Val(U4b_New_Combo_DestInt.Value) = 0 Then
        MsgBox ("Please Specify all Step/Interval Inputs!!")
        Exit Sub
    End If
    Worksheets("S6").Range("C11").Value = ""
    
    ' *** PART2 *** Add Stream Connection Info to Database
    ' [1] Add Stream Connection to Last Row on Sheet
    Dim Connections_Sheet As Worksheet
    Dim LastRowConnect As Integer
    Set Connections_Sheet = Sheets("B8")
    LastRowConnect = Connections_Sheet.Range("C65536").End(xlUp).Row
    LastRowConnect = LastRowConnect + 1

    ' [2] Update Connections List Sheet (B8) with new Connections from Userform
    Sheets("B8").Range("C" & LastRowConnect) = Me.U4b_New_Combo_SourceStep.Value
    Sheets("B8").Range("D" & LastRowConnect) = Me.U4b_New_Combo_SourceInt.Value
    Sheets("B8").Range("E" & LastRowConnect) = Me.U4b_New_Combo_DestStep.Value
    Sheets("B8").Range("F" & LastRowConnect) = Me.U4b_New_Combo_DestInt.Value
    
    ' [3] Delete Duplicate Entries
    Dim i As Long
    For i = 1 To Worksheets("B8").Cells.SpecialCells(xlLastCell).Row
        If Worksheets("B8").Cells(i, 3) <> vbNullString Then
            If Worksheets("B8").Cells(i, 3) = Worksheets("B8").Cells(i + 1, 3) And Worksheets("B8").Cells(i, 4) = Worksheets("B8").Cells(i + 1, 4) And Worksheets("B8").Cells(i, 5) = Worksheets("B8").Cells(i + 1, 5) And Worksheets("B8").Cells(i, 6) = Worksheets("B8").Cells(i + 1, 6) Then
               Worksheets("B8").Cells(i + 1, 1).EntireRow.Delete
                i = i - 1
            End If
        End If
    Next i
    
    
    ' *** PART3 *** Add 1 to Connectivity Matrix
    ' [1] Declare Variables
    Dim Isource As Integer
    Dim Idest As Integer
    
    ' [2] Find Row/Column Index *NOTE* Combobox returns entries as STRING NOT VALUES!!! A Conversion to Value is NECESSARY!!!
    total_intervals = Sheets("S4").Range("H14").Value + 7
    
    For Isource = 8 To total_intervals
        If Val(U4b_New_Combo_SourceStep.Value) = Worksheets("B7").Cells(Isource, 2).Value And Val(U4b_New_Combo_SourceInt.Value) = Worksheets("B7").Cells(Isource, 3).Value Then
        SourceInd = Isource
        End If
    Next Isource
    For Idest = 4 To total_intervals
        If Val(U4b_New_Combo_DestStep.Value) = Worksheets("B7").Cells(6, Idest).Value And Val(U4b_New_Combo_DestInt.Value) = Worksheets("B7").Cells(7, Idest).Value Then
        DestInd = Idest
        End If
    Next Idest
        
    ' [3] Assign a "1" to the cell with corresponding Row and Column Ind
    Worksheets("B7").Cells(SourceInd, DestInd).Value = 1
    
    ' *** PART4 *** ReDraw Network
    ' [1] Define Variables
    Dim Shp As Shape
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
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S4").Range("H12").Value + 2
    n_interval = Worksheets("S4").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value

    ' [3] Delete Results and Reset Checksums
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
       x = Worksheets("S4").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [7] Network Drawing Module
    x = 1
    For a = 1 To n_step
       ' Index steps and # of intervals in each step
       Current_Step = Worksheets("S4").Range("E" & 12 + a).Value
       current_interval = Worksheets("S4").Range("F" & 12 + a).Value
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
          If Worksheets("S4").Range("H12").Value >= 7 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
          End If
          If Worksheets("S4").Range("H12").Value >= 10 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
          End If
       
          'Locate shapes
          ' *** IF MAX INTERVAL IS 1 ***
          If max_interval = 1 Then
          Set Shp = ActiveSheet.Shapes("shape" & x)
          ' [NEW] Auto-Adjust Height between each Interval
          ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
          IntHeight = 450 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
          IntWidth = 1200 / ((2 * n_step) + (1.5 * (n_step - 1)))
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
      
    'Draw arrows to connect intevals
    'Primary connection
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

    'Secondary connection
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
    Range("A1").Select
End Sub

Private Sub U4b_Button_Add_S_Click()
    ' *** PART1 *** Check if all ComboBox Fields are Specified
    ' [1] Check that all Input Entries are Identified
    If Val(U4b_Sec_Combo_SourceStep.Value) = 0 Or Val(U4b_Sec_Combo_SourceInt.Value) = 0 Or Val(U4b_Sec_Combo_DestStep.Value) = 0 Or Val(U4b_Sec_Combo_DestInt.Value) = 0 Then
        MsgBox ("Please Specify all Step/Interval Inputs!!")
        Exit Sub
    End If
    Worksheets("S6").Range("C11").Value = ""
    
    ' *** PART2 *** Add Stream Connection Info to Database
    ' [1] Add Stream Connection to Last Row on Sheet
    Dim Connections_Sheet As Worksheet
    Dim LastRowConnect As Integer
    Set Connections_Sheet = Sheets("B9")
    LastRowConnect = Connections_Sheet.Range("C65536").End(xlUp).Row
    LastRowConnect = LastRowConnect + 1

    ' [2] Update Connections List Sheet (B8) with new Connections from Userform
    Sheets("B9").Range("C" & LastRowConnect) = Me.U4b_Sec_Combo_SourceStep.Value
    Sheets("B9").Range("D" & LastRowConnect) = Me.U4b_Sec_Combo_SourceInt.Value
    Sheets("B9").Range("E" & LastRowConnect) = Me.U4b_Sec_Combo_DestStep.Value
    Sheets("B9").Range("F" & LastRowConnect) = Me.U4b_Sec_Combo_DestInt.Value
    
    ' [3] Delete Duplicate Entries
    Dim i As Long
    For i = 1 To Worksheets("B9").Cells.SpecialCells(xlLastCell).Row
        If Worksheets("B9").Cells(i, 3) <> vbNullString Then
            If Worksheets("B9").Cells(i, 3) = Worksheets("B9").Cells(i + 1, 3) And Worksheets("B9").Cells(i, 4) = Worksheets("B9").Cells(i + 1, 4) And Worksheets("B9").Cells(i, 5) = Worksheets("B9").Cells(i + 1, 5) And Worksheets("B9").Cells(i, 6) = Worksheets("B9").Cells(i + 1, 6) Then
               Worksheets("B9").Cells(i + 1, 1).EntireRow.Delete
                i = i - 1
            End If
        End If
    Next i
    
    
    ' *** PART3 *** Add 1 to Connectivity Matrix
    ' [1] Declare Variables
    Dim Isource As Integer
    Dim Idest As Integer
    
    ' [2] Find Row/Column Index *NOTE* Combobox returns entries as STRING NOT VALUES!!! A Conversion to Value is NECESSARY!!!
    total_intervals = Sheets("S4").Range("H14").Value + 13
    total_intervals2 = (2 * Sheets("S4").Range("H14").Value) + 12
    
    For Isource = total_intervals To total_intervals2
        If Val(U4b_Sec_Combo_SourceStep.Value) = Worksheets("B7").Cells(Isource, 2).Value And Val(U4b_Sec_Combo_SourceInt.Value) = Worksheets("B7").Cells(Isource, 3).Value Then
        SourceInd = Isource
        End If
    Next Isource
    For Idest = 4 To total_intervals
        If Val(U4b_Sec_Combo_DestStep.Value) = Worksheets("B7").Cells(total_intervals - 2, Idest).Value And Val(U4b_Sec_Combo_DestInt.Value) = Worksheets("B7").Cells(total_intervals - 1, Idest).Value Then
        DestInd = Idest
        End If
    Next Idest
        
    ' [3] Assign a "1" to the cell with corresponding Row and Column Ind
    Worksheets("B7").Cells(SourceInd, DestInd).Value = 1
    
    ' *** PART4 *** ReDraw Network
    ' [1] Define Variables
    Dim Shp As Shape
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
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S4").Range("H12").Value + 2
    n_interval = Worksheets("S4").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value

    ' [3] Delete Results and Reset Checksums
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
       x = Worksheets("S4").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [7] Network Drawing Module
    x = 1
    For a = 1 To n_step
       ' Index steps and # of intervals in each step
       Current_Step = Worksheets("S4").Range("E" & 12 + a).Value
       current_interval = Worksheets("S4").Range("F" & 12 + a).Value
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
          If Worksheets("S4").Range("H12").Value >= 7 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
          End If
          If Worksheets("S4").Range("H12").Value >= 10 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
          End If
       
          'Locate shapes
          ' *** IF MAX INTERVAL IS 1 ***
          If max_interval = 1 Then
          Set Shp = ActiveSheet.Shapes("shape" & x)
          ' [NEW] Auto-Adjust Height between each Interval
          ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
          IntHeight = 450 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
          IntWidth = 1200 / ((2 * n_step) + (1.5 * (n_step - 1)))
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
      
    'Draw arrows to connect intevals
    'Primary connection
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

    'Secondary connection
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
    Range("A1").Select
End Sub



' Removing a PSIN Stream Connection then Redrawing Network
Private Sub U4b_RemoveConnection_Click()
' [1] Declare Variables
Dim msg As String
Dim ans As String
Dim rowselect As Long
Dim ConNameSStep As Long
Dim ConNameSInt As Long
Dim ConNameDStep As Long
Dim ConNameDInt As Long
Dim SourceStep As Long
Dim SourceInt As Long
Dim DestStep As Long
Dim DestInt As Long

' [2] Must Select Connection Before Proceeding
If Me.U4b_ConnectionsList.Value = "" Then
MsgBox "Please select a Stream Connection to Remove!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If
Worksheets("S6").Range("C11").Value = ""
  
' [3] Connection Name as String.
ConNameSStep = Val(U4b_ConnectionsList.Column(0))
ConNameSInt = Val(U4b_ConnectionsList.Column(1))
ConNameDStep = Val(U4b_ConnectionsList.Column(2))
ConNameDInt = Val(U4b_ConnectionsList.Column(3))

' [4] Use ConFind function to find Row Index, also remove the row with the connectivity values from datasheet B8
rowselect = ConFind(ConNameSStep, ConNameSInt, ConNameDStep, ConNameDInt)
    SourceStep = Sheets("B8").Cells(rowselect, 3).Value
    SourceInt = Sheets("B8").Cells(rowselect, 4).Value
    DestStep = Sheets("B8").Cells(rowselect, 5).Value
    DestInt = Sheets("B8").Cells(rowselect, 6).Value
Sheets("B8").Rows(rowselect).EntireRow.Delete
rowselect = rowselect - 1

' [5] Find the "1" in the Connectivity Matrix
' Find Row/Column Index *NOTE* Combobox returns entries as STRING NOT VALUES!!! A Conversion to Value is NECESSARY!!!
total_intervals = Sheets("S4").Range("H14").Value + 7
Dim Isource As Integer
Dim Idest As Integer
For Isource = 8 To total_intervals
    If Val(SourceStep) = Worksheets("B7").Cells(Isource, 2).Value And Val(SourceInt) = Worksheets("B7").Cells(Isource, 3).Value Then
    SourceInd = Isource
    End If
Next Isource
For Idest = 4 To total_intervals
    If Val(DestStep) = Worksheets("B7").Cells(6, Idest).Value And Val(DestInt) = Worksheets("B7").Cells(7, Idest).Value Then
    DestInd = Idest
    End If
Next Idest
Worksheets("B7").Cells(SourceInd, DestInd).Value = None

' [6] Update Process Network
    ' ***PROCESS NETWORK DRAWING MODULE*** '
    ' [6-1] Define Variables
    Dim Shp As Shape
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
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [6-2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S4").Range("H12").Value + 2
    n_interval = Worksheets("S4").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value

    ' [6-3] Delete Results and Reset Checksums
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
    
    ' [6-3] Change Variable Dimension to n_int by n_int Array
    ReDim connection_p(1 To n_interval, 1 To n_interval)
    ReDim connection_s(1 To n_interval, 1 To n_interval)
    
    ' [6-4] Delete any previous Networks
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
       
    ' [6-5] Recall Connectivity Information
    For a = 1 To n_interval
       For b = 1 To n_interval
          connection_p(a, b) = Worksheets("B7").Cells(7 + a, 3 + b).Value
          connection_s(a, b) = Worksheets("B7").Cells(12 + n_interval + a, 3 + b).Value
       Next b
    Next a
    
    ' [6-6] Find maximum processing interval number
    max_interval = 0
    For a = 1 To n_step
       x = Worksheets("S4").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [6-7] Network Drawing Module
    x = 1
    For a = 1 To n_step
       ' Index steps and # of intervals in each step
       Current_Step = Worksheets("S4").Range("E" & 12 + a).Value
       current_interval = Worksheets("S4").Range("F" & 12 + a).Value
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
          If Worksheets("S4").Range("H12").Value >= 7 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
          End If
          If Worksheets("S4").Range("H12").Value >= 10 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
          End If
       
          'Locate shapes
          ' *** IF MAX INTERVAL IS 1 ***
          If max_interval = 1 Then
          Set Shp = ActiveSheet.Shapes("shape" & x)
          ' [NEW] Auto-Adjust Height between each Interval
          ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
          IntHeight = 450 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
          IntWidth = 1200 / ((2 * n_step) + (1.5 * (n_step - 1)))
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
      
    ' [6-8] Draw arrows to connect intevals
    'Primary connection
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

    'Secondary connection
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
    Range("A1").Select

' [7] Show Msg that Material Was Deleted
ans = MsgBox("Connection has been deleted", vbOKOnly, "TIPEM- Stream Connection Deleted")
If ans = vbYes Then
    U4b_Stream_Connections.Show
End If
End Sub

Private Sub U4b_RemoveConnection_S_Click()
' [1] Declare Variables
Dim msg As String
Dim ans As String
Dim rowselect2 As Long
Dim ConNameSStep2 As Long
Dim ConNameSInt2 As Long
Dim ConNameDStep2 As Long
Dim ConNameDInt2 As Long
Dim SourceStep2 As Long
Dim SourceInt2 As Long
Dim DestStep2 As Long
Dim DestInt2 As Long

' [2] Must Select Connection Before Proceeding
If Me.U4b_ConnectionsList_S.Value = "" Then
MsgBox "Please select a Stream Connection to Remove!!", vbExclamation, "TIPEM- Error"
Exit Sub
End If
Worksheets("S6").Range("C11").Value = ""
  
' [3] Connection Name as String.
ConNameSStep2 = Val(U4b_ConnectionsList_S.Column(0))
ConNameSInt2 = Val(U4b_ConnectionsList_S.Column(1))
ConNameDStep2 = Val(U4b_ConnectionsList_S.Column(2))
ConNameDInt2 = Val(U4b_ConnectionsList_S.Column(3))

' [4] Use ConFind function to find Row Index, also remove the row with the connectivity values from datasheet B8
rowselect2 = ConFind2(ConNameSStep2, ConNameSInt2, ConNameDStep2, ConNameDInt2)
    SourceStep2 = Sheets("B9").Cells(rowselect2, 3).Value
    SourceInt2 = Sheets("B9").Cells(rowselect2, 4).Value
    DestStep2 = Sheets("B9").Cells(rowselect2, 5).Value
    DestInt2 = Sheets("B9").Cells(rowselect2, 6).Value
Sheets("B9").Rows(rowselect2).EntireRow.Delete
rowselect2 = rowselect2 - 1

' [5] Find the "1" in the Connectivity Matrix
' Find Row/Column Index *NOTE* Combobox returns entries as STRING NOT VALUES!!! A Conversion to Value is NECESSARY!!!
    total_intervals = Sheets("S4").Range("H14").Value + 13
    total_intervals2 = (2 * Sheets("S4").Range("H14").Value) + 12
Dim Isource As Integer
Dim Idest As Integer
For Isource = total_intervals To total_intervals2
    If Val(SourceStep2) = Worksheets("B7").Cells(Isource, 2).Value And Val(SourceInt2) = Worksheets("B7").Cells(Isource, 3).Value Then
    SourceInd2 = Isource
    End If
Next Isource
For Idest = 4 To total_intervals
    If Val(DestStep2) = Worksheets("B7").Cells(total_intervals - 2, Idest).Value And Val(DestInt2) = Worksheets("B7").Cells(total_intervals - 1, Idest).Value Then
    DestInd2 = Idest
    End If
Next Idest
Worksheets("B7").Cells(SourceInd2, DestInd2).Value = None

' [6] Update Process Network
    ' ***PROCESS NETWORK DRAWING MODULE*** '
    ' [6-1] Define Variables
    Dim Shp As Shape
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
    Dim IntHeight As Single
    Dim IntWidth As Single
    
    ' [6-2] Assign Total Step and Interval Numbers
    n_step = Worksheets("S4").Range("H12").Value + 2
    n_interval = Worksheets("S4").Range("H14").Value
    n_mat = Worksheets("B2").Range("K3").Value

    ' [6-3] Delete Results and Reset Checksums
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
    
    ' [6-3] Change Variable Dimension to n_int by n_int Array
    ReDim connection_p(1 To n_interval, 1 To n_interval)
    ReDim connection_s(1 To n_interval, 1 To n_interval)
    
    ' [6-4] Delete any previous Networks
    For Each Shp In ActiveSheet.Shapes
       If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
       
    ' [6-5] Recall Connectivity Information
    For a = 1 To n_interval
       For b = 1 To n_interval
          connection_p(a, b) = Worksheets("B7").Cells(7 + a, 3 + b).Value
          connection_s(a, b) = Worksheets("B7").Cells(12 + n_interval + a, 3 + b).Value
       Next b
    Next a
    
    ' [6-6] Find maximum processing interval number
    max_interval = 0
    For a = 1 To n_step
       x = Worksheets("S4").Range("F" & 12 + a).Value
       If x > max_interval Then
          max_interval = x
       End If
    Next a
    
    ' [6-7] Network Drawing Module
    x = 1
    For a = 1 To n_step
       ' Index steps and # of intervals in each step
       Current_Step = Worksheets("S4").Range("E" & 12 + a).Value
       current_interval = Worksheets("S4").Range("F" & 12 + a).Value
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
          If Worksheets("S4").Range("H12").Value >= 7 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 9
          End If
          If Worksheets("S4").Range("H12").Value >= 10 Then
             Selection.ShapeRange(1).TextFrame2.TextRange.Font.Size = 7
          End If
       
          'Locate shapes
          ' *** IF MAX INTERVAL IS 1 ***
          If max_interval = 1 Then
          Set Shp = ActiveSheet.Shapes("shape" & x)
          ' [NEW] Auto-Adjust Height between each Interval
          ' TOTAL AVAILABLE HEIGHT IS 300 PIXELS. TOTAL AVAILABLE WIDTH IS 820 PIXELS
          IntHeight = 450 / ((2 * max_interval) + (1.5 * (max_interval - 1)))
          IntWidth = 1200 / ((2 * n_step) + (1.5 * (n_step - 1)))
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
      
    ' [6-8] Draw arrows to connect intevals
    'Primary connection
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

    'Secondary connection
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
    Range("A1").Select

' [7] Show Msg that Material Was Deleted
ans = MsgBox("Connection has been deleted", vbOKOnly, "TIPEM- Stream Connection Deleted")
If ans = vbYes Then
    U4b_Stream_Connections.Show
End If
End Sub



' Cancel Userform by Closing Userform
Private Sub U4b_Button_Cancel_Click()
Unload Me
End
End Sub
Private Sub U4b_Button_Cancel_S_Click()
Unload Me
End
End Sub



' Ok closes the Userform
Private Sub U4b_Button_Ok_Click()
Unload Me
End Sub
Private Sub U4b_Button_Ok_S_Click()
Unload Me
End Sub
