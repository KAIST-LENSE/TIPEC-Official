VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5g_AssignIntervalNames 
   Caption         =   "Assign Names to Intervals"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7200
   OleObjectBlob   =   "U5g_AssignIntervalNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5g_AssignIntervalNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub U5g_Cancel_Click()
Unload Me
End
End Sub


Private Sub U5g_ComboBox1_Change()
    ' ComboBox to Select Intervals
    Dim Num_Steps As Integer
    Num_Steps = Worksheets("S4").Cells(12, 8).Value + 2
    
    ' If Step = 1, AKA Feedstock Step, then:
    If U5g_ComboBox1.Column(0) = 1 Then
        U5g_DisplayInterval = "Feedstock" & "-" & U5g_ComboBox1.Column(1) & "   |   " & U5g_ComboBox1.Column(2)
    ' If Step = Last Step, AKA Product Step, then:
    ElseIf U5g_ComboBox1.Column(0) = Num_Steps Then
        U5g_DisplayInterval = "Product" & "-" & U5g_ComboBox1.Column(1) & "   |   " & U5g_ComboBox1.Column(2)
    ' Else, it is a process step
    Else
        U5g_DisplayInterval = "Process Step " & U5g_ComboBox1.Column(0) & "-" & U5g_ComboBox1.Column(1) & "   |   " & U5g_ComboBox1.Column(2)
    End If
End Sub


Private Sub U5g_Save_Click()
    ' Make sure text is entered in the field
    If Me.U5g_IntName.Value = "" Then
    MsgBox "Please Enter a Name for the selected Interval", vbExclamation, "TIPEM- Error"
    Exit Sub
    End If
    
    ' Find maximum number of intervals
    Dim row_source As Integer
    Num_Int = Sheets("S4").Range("H14").Value + 7

    ' Loop from B8 to B8+NumInt, Find the Row Matching Step/Interval Index
    For a = 8 To Num_Int
        If U5g_ComboBox1.Column(0) = Worksheets("B10").Cells(a, 2).Value And U5g_ComboBox1.Column(1) = Worksheets("B10").Cells(a, 3).Value Then
            row_source = a
        End If
    Next a
    Worksheets("B10").Cells(row_source, 4).Value = U5g_IntName.Value
End Sub

Private Sub UserForm_Click()

End Sub
