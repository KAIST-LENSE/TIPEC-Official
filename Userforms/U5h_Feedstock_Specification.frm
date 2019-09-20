VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5h_Feedstock_Specification 
   Caption         =   "Feedstock Specification"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   OleObjectBlob   =   "U5h_Feedstock_Specification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5h_Feedstock_Specification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub U5h_Exit_Click()
Unload Me
End
End Sub




Private Sub U5h_MassFrac_Change()
On Error Resume Next
' [1] Save the Mass Fraction
    ' Find the current row corresponding to the Feed Interval Selected
    Dim ind_start As Integer            ' The starting row index for the tables, which is currently 8
    Dim ind_intervals As Integer        ' The number of rows corresponding to the number of intervals
    Dim ind_space As Integer            ' The space between each table, corresponding to 4
    Dim ind_heading As Integer          ' The number of rows taken by headings
    Dim array_start As Integer
    Dim array_end As Integer
    Dim Num_Mat As Integer
    Dim RowFound As Integer
    Dim ColFound As Integer
    
    ' Assign Indices
    ind_start = 8
    ind_intervals = Sheets("S4").Range("H14").Value
    ind_space = 4
    ind_heading = 2
    Num_Mat = Worksheets("B2").Range("K3").Value
    array_start = ind_start + ind_intervals + ind_space + ind_heading
    array_end = array_start + Sheets("S4").Range("F13").Value
    
    ' Find row corresponding to selected interval
    For a = array_start To array_end
        If Worksheets("B10").Range("H3").Value = Worksheets("B10").Cells(a, 2).Value And Worksheets("B10").Range("K3").Value = Worksheets("B10").Cells(a, 3).Value Then
            RowFound = a
        End If
    Next a
    
    ' Find the COLUMN corresponding to the material
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(array_start - 1, 3 + b).Value = Me.U5h_MaterialsList.Column(1) Then
            ColFound = 3 + b
        End If
    Next b

    ' Assign Mass Fraction to the Found Cell
    Worksheets("B10").Cells(RowFound, ColFound).Value = Me.U5h_MassFrac.Value

    ' Update Display of Total Mass Frac
    Me.U5h_Label6 = Application.Sum(Range(Worksheets("B10").Cells(RowFound, 4), Worksheets("B10").Cells(RowFound, 3 + Num_Mat)))

End Sub




Private Sub U5h_Save_Click()
' [1] Save the Feedrate
    ' Find the current row corresponding to the Feed Interval Selected
    Dim ind_start As Integer            ' The starting row index for the tables, which is currently 8
    Dim ind_intervals As Integer        ' The number of rows corresponding to the number of intervals
    Dim ind_space As Integer            ' The space between each table, corresponding to 4
    Dim ind_heading As Integer          ' The number of rows taken by headings
    Dim array_start As Integer
    Dim array_end As Integer
    Dim Num_Mat As Integer
    Dim RowFound As Integer
    Dim ColFound As Integer
    
    ' Assign Indices
    ind_start = 8
    ind_intervals = Sheets("S4").Range("H14").Value
    ind_space = 4
    ind_heading = 2
    Num_Mat = Worksheets("B2").Range("K3").Value
    array_start = ind_start + ind_intervals + ind_space + ind_heading
    array_end = array_start + Sheets("S4").Range("F13").Value
    
    ' Find row corresponding to selected interval
    For a = array_start To array_end
        If Worksheets("B10").Range("H3").Value = Worksheets("B10").Cells(a, 2).Value And Worksheets("B10").Range("K3").Value = Worksheets("B10").Cells(a, 3).Value Then
            RowFound = a
        End If
    Next a
    
    ' After finding row, save the current feedrate
    If U5h_Feedrate.Value = "" Then
        Worksheets("B10").Cells(RowFound, 4 + Num_Mat).Value = 0
    Else
        Worksheets("B10").Cells(RowFound, 4 + Num_Mat).Value = U5h_Feedrate.Value
    End If
    
    
' [2] Save Feedrate Mass Basis
    ' Loop through indices in B1, and find one with "1" in it
    For b = 27 To 29
        If Worksheets("B1").Cells(b, 2).Value = U5h_BasisUnits.Value Then
            Worksheets("B1").Cells(b, 3).Value = 1
        Else
            Worksheets("B1").Cells(b, 3).Value = 0
        End If
    Next b
    
    
' [3] Create a Custom Array with Materials and Feedstock Composition
    Dim FEED_array()
    ReDim FEED_array(Num_Mat, 3)

    For i = 1 To Num_Mat
        FEED_array(i - 1, 0) = Worksheets("B2").Cells(3 + i, 2).Value
        FEED_array(i - 1, 1) = Worksheets("B2").Cells(3 + i, 3).Value
        If Worksheets("B10").Cells(RowFound, 3 + i).Value = "" Then
            FEED_array(i - 1, 2) = 0
        Else
            FEED_array(i - 1, 2) = Worksheets("B10").Cells(RowFound, 3 + i).Value
        End If
    Next i
    Me.U5h_MaterialsList.list = FEED_array
' [3] Check if Sum of Feedstock Mass Fractions is not 1
    'If Application.Sum(Range(Worksheets("B10").Cells(RowFound, 4), Worksheets("B10").Cells(RowFound, 3 + Num_Mat))) <> 1 Then
    '    MsgBox "The sum of material mass fractions DOES NOT equal 1!! Try Again!!", vbExclamation, "Error"
    'End If
End Sub




Private Sub UserForm_Initialize()
' [1] Load Feedrate
    ' Find the ROW corresponding to the Feed Interval Selected
    Dim ind_intervals As Integer
    Dim array_start As Integer
    Dim Num_Feed As Integer
    Dim Num_Mat As Integer
    Dim RowFound As Integer
    Dim ColFound As Integer
    
    ' Assign Indices
    ind_intervals = Worksheets("S4").Range("H14").Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    Num_Feed = Worksheets("S4").Range("F13").Value
    array_start = 7 + ind_intervals + 6
    
    ' Find row corresponding to selected interval
    For a = 1 To Num_Feed
        If Worksheets("B10").Range("H3").Value = Worksheets("B10").Cells(array_start + a, 2).Value And Worksheets("B10").Range("K3").Value = Worksheets("B10").Cells(array_start + a, 3).Value Then
            RowFound = array_start + a
        End If
    Next a
    
    ' After finding row, display the current feedrate (Default = Empty)
    U5h_Feedrate.Value = Worksheets("B10").Cells(RowFound, 4 + Num_Mat).Value


' [2] Load Feedrate Mass Basis
    ' Loop through indices in B1, and find one with "1" in it
    For b = 23 To 25
        If Worksheets("B1").Cells(b, 3).Value = 1 Then
            U5h_BasisUnits.Value = Worksheets("B1").Cells(b, 2).Value
        Else
            U5h_BasisUnits.Value = Worksheets("B1").Cells(23, 2).Value
        End If
    Next b
    
    
' [3] Create a Custom Array with Materials and Feedstock Composition
    Dim FEED_array()
    ReDim FEED_array(Num_Mat, 3)

    For i = 1 To Num_Mat
        FEED_array(i - 1, 0) = Worksheets("B2").Cells(3 + i, 2).Value
        FEED_array(i - 1, 1) = Worksheets("B2").Cells(3 + i, 3).Value
        If Worksheets("B10").Cells(RowFound, 3 + i).Value = "" Then
            FEED_array(i - 1, 2) = 0
            Worksheets("B10").Cells(RowFound, 3 + i).Value = 0
        Else
            FEED_array(i - 1, 2) = Worksheets("B10").Cells(RowFound, 3 + i).Value
        End If
    Next i
    Me.U5h_MaterialsList.list = FEED_array
End Sub




Private Sub U5h_MaterialsList_Change()
' [1] Display Selected Material Info
    U5h_DisplayMaterial.Text = U5h_MaterialsList.Column(0) & "   |   " & U5h_MaterialsList.Column(1)

' [2] Material Selected from List
    If Me.U5h_MaterialsList.Value = "" Then
        MsgBox "Please select a Material to specify its composition!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If

' [3] Load Compositional Values
    ' Find the ROW corresponding to the Feed Interval Selected
    Dim ind_start As Integer            ' The starting row index for the tables, which is currently 8
    Dim ind_intervals As Integer        ' The number of rows corresponding to the number of intervals
    Dim ind_space As Integer            ' The space between each table, corresponding to 4
    Dim ind_heading As Integer          ' The number of rows taken by headings
    Dim array_start As Integer
    Dim array_end As Integer
    Dim Num_Mat As Integer
    Dim RowFound As Integer
    Dim ColFound As Integer
    
    ' Assign Indices
    ind_start = 8
    ind_intervals = Sheets("S4").Range("H14").Value
    ind_space = 4
    ind_heading = 2
    Num_Mat = Worksheets("B2").Range("K3").Value
    array_start = ind_start + ind_intervals + ind_space + ind_heading
    array_end = array_start + Sheets("S4").Range("F13").Value
    
    ' Find row corresponding to selected interval
    For a = array_start To array_end
        If Worksheets("B10").Range("H3").Value = Worksheets("B10").Cells(a, 2).Value And Worksheets("B10").Range("K3").Value = Worksheets("B10").Cells(a, 3).Value Then
            RowFound = a
        End If
    Next a
    
    ' Find the COLUMN corresponding to the material
    For b = 1 To Num_Mat
        If Worksheets("B10").Cells(array_start - 1, 3 + b).Value = Me.U5h_MaterialsList.Column(1) Then
            ColFound = 3 + b
        End If
    Next b

    ' Assign Mass Fraction to the Found Cell
    Me.U5h_MassFrac.Value = Worksheets("B10").Cells(RowFound, ColFound).Value
End Sub
