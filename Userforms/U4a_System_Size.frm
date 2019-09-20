VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U4a_System_Size 
   Caption         =   "Specify the Superstructure System Size"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7470
   OleObjectBlob   =   "U4a_System_Size.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U4a_System_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U4a_Button_Ok_Click()
    '***********************************************************
    '************ DELETE CONNECTIVITY MATRIX *******************
    '***********************************************************
    ' [1] Define Variables
    Application.ScreenUpdating = False
    Application.Goto (Sheets("B7").Range("B4:CZ220"))
    Worksheets("B7").Range("B4:CZ220").ClearContents
    Selection.Font.Bold = False
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Interior.TintAndShade = 0
    Dim Shp As Shape
    For Each Shp In ActiveSheet.Shapes
        If Not (Shp.Type = msoOLEControlObject Or Shp.Type = msoFormControl) Then Shp.Delete
    Next Shp
    Application.ScreenUpdating = True
    Worksheets("S4").Activate
    
    
    '***********************************************************
    '************ GENERATE STEP/INTERVAL TABLE *****************
    '***********************************************************
    ' [1] Declare Variables
    Dim input_cell As Range
    Dim aaa As Integer
    Dim nnn As Integer
    Dim st As Integer
    
    ' [2] If no Step # Value is Entered, then Close Sub
    If Len(U4a_Input1.Value) = 0 Then Exit Sub
    Set input_cell = Cells(12, 8)
      
    ' [3] Store Number of Steps in input_cell
    With input_cell
       .Value = Val(Replace(U4a_Input1.Value, ",", ""))
       .NumberFormat = "#,###"
    End With
    st = U4a_Input1.Value
    
    ' [4] Generate Feedstock Intervals
    Range("D" & 13).Value = "Feedstock Int."
    Range("E" & 13).Value = 1
    Range("F" & 13).Value = U4a_Feedstock.Value
    Range("D13:F13").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
    
    ' [5] Generate Number of Steps for Entering Intervals
    For aaa = 1 To st
        Range("D" & 13 + aaa).Value = "Process Step " & aaa
        Range("E" & 13 + aaa).Value = aaa + 1
        Range("F" & 13 + aaa).Value = "Enter Interval #"
        With Range("F" & 13 + aaa).Font
            .Color = -16776961
            .Italic = True
        End With
        With Range("E" & 13 + aaa).Font
            .Color = -16776961
            .Italic = True
        End With
        With Range("D" & 13 + aaa).Font
            .Color = -16776961
            .Italic = True
        End With
    Next aaa
    
    ' [6] Generate Product Inverals
    Range("D" & 14 + st).Value = "Product Int."
    Range("E" & 14 + st).Value = st + 2
    Range("F" & 14 + st).Value = U4a_Products.Value
    Range("D" & 14 + st, "F" & 14 + st).Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.249977111117893
    End With
    
    ' [6] Color Step/Interval Table by Selecting the End Index of Table: 1st and 2nd Columns
    Range("D" & 14 + Range("H12").Value, "F" & 14 + Range("H12").Value).Select
    With Selection
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
    End With
    With Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .ColorIndex = 0
         .TintAndShade = 0
         .Weight = xlThin
    End With

    ' [7] Color Step/Interval Table by Selecting the End Index of Table: Interval Column
    Range("F" & 14, "F" & 13 + Range("H12").Value).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
      ' [8] Center Text
    Range("D12", "E" & 12 + aaa).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Range("F12", "F" & 12 + aaa).Select
        With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Range("A1").Select
    Unload Me
End Sub


Private Sub U4a_Button_Cancel_Click()
   Unload Me
End Sub


Private Sub U4a_Input1_Change()
   With U4a_Input1
      .Text = Format(.Value, "#,###")
      .TextAlign = fmTextAlignCenter
   End With
End Sub


Private Sub UserForm_Initialize()
   ' [1] Declare Variables
   Dim n As Integer
   
   ' [2] Change Color and Style of PSIN Table
   n = Range("H12").Value
   Range("D" & 12 + n, "F" & 12).Select
   With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' [3] Center Text
    Range("D12", "E" & 12 + n).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Range("F13", "F" & 13 + n).Select
        With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    End With
    Range("A1").Select
End Sub

Private Sub U4a_Input1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then KeyAscii = 0
End Sub

