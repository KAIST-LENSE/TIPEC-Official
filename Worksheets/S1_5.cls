VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***UPDATING THE TO-DO LIST***
Private Sub Worksheet_Activate()
    '(1) Check if Mass Balances are Available
    If Worksheets("O1").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 17")).Select
        With Selection
            .Text = "¡î  -  Process Specification is complete."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 17")).Select
        With Selection
            .Text = "X  -  Process Specification is complete."
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
    
    '(2) Check if Plant Results Summary is Available
    If Worksheets("O2").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 20")).Select
        With Selection
            .Text = "¡î  -  Mass Balances are available."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 20")).Select
        With Selection
            .Text = "X  -  Mass Balances are available."
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

    '(3) Check if Equipment Costs are Available for all Intervals
    If Worksheets("O3").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 21")).Select
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
        Worksheets("S5").Shapes.Range(Array("TextBox 21")).Select
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

    '(4) Check if Capital Costs are Available
    If Worksheets("O4").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 22")).Select
        With Selection
            .Text = "¡î  -  CAPEX Lang Factors have been specified."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 22")).Select
        With Selection
            .Text = "X  -  CAPEX Lang Factors have been specified."
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

    '(5) Check if Operating Costs are Available
    If Worksheets("O4").Range("H2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 23")).Select
        With Selection
            .Text = "¡î  -  OPEX Lang Factors have been specified."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 23")).Select
        With Selection
            .Text = "X  -  OPEX Lang Factors have been specified."
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
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error Resume Next
    '(1) Check if Mass Balances are Available
    If Worksheets("O1").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 17")).Select
        With Selection
            .Text = "¡î  -  Process Specification is complete."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 17")).Select
        With Selection
            .Text = "X  -  Process Specification is complete."
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
    
    '(2) Check if Plant Results Summary is Available
    If Worksheets("O2").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 20")).Select
        With Selection
            .Text = "¡î  -  Mass Balances are available."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 20")).Select
        With Selection
            .Text = "X  -  Mass Balances are available."
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

    '(3) Check if Equipment Costs are Available for all Intervals
    If Worksheets("O3").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 21")).Select
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
        Worksheets("S5").Shapes.Range(Array("TextBox 21")).Select
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

    '(4) Check if Capital Costs are Available
    If Worksheets("O4").Range("F2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 22")).Select
        With Selection
            .Text = "¡î  -  CAPEX Lang Factors have been specified."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 22")).Select
        With Selection
            .Text = "X  -  CAPEX Lang Factors have been specified."
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

    '(5) Check if Operating Costs are Available
    If Worksheets("O4").Range("H2").Value = 1 Then
        Worksheets("S5").Shapes.Range(Array("TextBox 23")).Select
        With Selection
            .Text = "¡î  -  OPEX Lang Factors have been specified."
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
        Worksheets("S5").Shapes.Range(Array("TextBox 23")).Select
        With Selection
            .Text = "X  -  OPEX Lang Factors have been specified."
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
End Sub

