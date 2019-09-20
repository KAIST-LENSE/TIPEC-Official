VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'***REDRAWING THE SELECTED PATHWAY NETWORK***
Private Sub Worksheet_Activate()
    For Each Shp2 In Worksheets("S8").Shapes
       If Not (Shp2.Type = msoOLEControlObject Or Shp2.Type = msoFormControl) Then Shp2.Delete
    Next Shp2

    Network_Drawing_Pathway_Config
    Range("A1").Select
    
    ' Display Selected Feedstock Interval
    Num_Feed = Worksheets("S3").Range("F13").Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    
    For i = 1 To Num_Feed
        For j = 1 To Num_Mat
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                Worksheets("S8").Range("F12").Value = i
                Worksheets("S8").Range("G12").Value = Worksheets("B10").Cells(7 + i, 4).Value
            End If
        Next j
    Next i
End Sub


'***ENTERING IN THE INDEX OF THE KEY PRODUCT INTERVAL***
'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim KeyCells As Range
'    Set KeyCells = Worksheets("S6").Range("I12")
'    n_prod = Worksheets("S4").Cells(Rows.Count, "F").End(xlUp).Value
'
'    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
'        If KeyCells.Value = "" Then
'            MsgBox "Please enter a valid index of the Key Product Interval!!!", vbExclamation, "Error"
'            KeyCells.Value = 1
'            Exit Sub
'        End If
'        If KeyCells.Value < 1 Or KeyCells.Value > n_prod Then
'            MsgBox "Please enter a valid index of the Key Product Interval!!!", vbExclamation, "Error"
'            KeyCells.Value = 1
'            Exit Sub
'        End If
'    End If
'End Sub

