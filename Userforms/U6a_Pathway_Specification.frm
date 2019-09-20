VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U6a_Pathway_Specification 
   Caption         =   "Define Stream Connections"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   OleObjectBlob   =   "U6a_Pathway_Specification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U6a_Pathway_Specification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MultiPage1_Change()

End Sub

'***UPDATE WHAT FEEDSTOCK IS CHOSEN***
Private Sub UserForm_Deactivate()
    Num_Feed = Worksheets("S4").Range("F13").Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    
    For i = 1 To Num_Feed
        For j = 1 To Num_Mat
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                Worksheets("S6").Range("F12").Value = i
                Worksheets("S6").Range("G12").Value = Worksheets("B10").Cells(7 + i, 4).Value
            End If
        Next j
    Next i
End Sub
Private Sub UserForm_Terminate()
    Num_Feed = Worksheets("S4").Range("F13").Value
    Num_Mat = Worksheets("B2").Range("K3").Value
    
    For i = 1 To Num_Feed
        For j = 1 To Num_Mat
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                Worksheets("S6").Range("F12").Value = i
                Worksheets("S6").Range("G12").Value = Worksheets("B10").Cells(7 + i, 4).Value
            End If
        Next j
    Next i
End Sub





'***USERFORM INITIALIZE***
Private Sub UserForm_Initialize()
On Error Resume Next
' [1] Create a Custom Array to display for Current Pathway Combobox
    'Count the number of chosen connections
    total_intervals = Sheets("S4").Range("H14").Value
    ConCount = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(8, 4), Worksheets("B12").Cells(8 + total_intervals, 4 + total_intervals)))
    'Define an array with list of added connections.
    Dim PC_Chosen()
    Dim k As Integer
    ReDim PC_Chosen(ConCount, 4)
    
    k = 1
    For i = 1 To total_intervals
        For j = 1 To total_intervals
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                PC_Chosen(k - 1, 0) = Worksheets("B12").Cells(7 + i, 2).Value
                PC_Chosen(k - 1, 1) = Worksheets("B12").Cells(7 + i, 3).Value
                PC_Chosen(k - 1, 2) = Worksheets("B12").Cells(6, 3 + j).Value
                PC_Chosen(k - 1, 3) = Worksheets("B12").Cells(7, 3 + j).Value
                k = k + 1
            End If
        Next j
    Next i
    Me.U6a_PCSelected_Combobox.list = PC_Chosen
    
    'Count the number of chosen connections
    ConCount2 = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(13 + total_intervals, 4), Worksheets("B12").Cells(13 + total_intervals + total_intervals, 4 + total_intervals)))
    Dim SC_Chosen()
    Dim kk As Integer
    ReDim SC_Chosen(ConCount2, 4)
    
    kk = 1
    For ii = 1 To total_intervals
        For jj = 1 To total_intervals
            If Worksheets("B12").Cells(12 + total_intervals + ii, 3 + jj).Value = 1 Then
                SC_Chosen(kk - 1, 0) = Worksheets("B12").Cells(12 + total_intervals + ii, 2).Value
                SC_Chosen(kk - 1, 1) = Worksheets("B12").Cells(12 + total_intervals + ii, 3).Value
                SC_Chosen(kk - 1, 2) = Worksheets("B12").Cells(11 + total_intervals, 3 + jj).Value
                SC_Chosen(kk - 1, 3) = Worksheets("B12").Cells(12 + total_intervals, 3 + jj).Value
                kk = kk + 1
            End If
        Next jj
    Next ii
    Me.U6a_SCSelected_Combobox.list = SC_Chosen

' [2] Export Number of Connections
    Worksheets("B12").Range("H2").Value = ConCount
    Worksheets("B12").Range("J2").Value = ConCount2
End Sub





'***ADD CONNECTION TO CURRENTLY SELECTED LIST***
Private Sub U6a_PC_Add_Click()
On Error Resume Next
' [1] Make sure a connection has been selected
    If Me.U6a_PC_Available_Combobox.Value = "" Then
        MsgBox "Please select a Primary Stream Connection to add to current pathway!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Get the Indexes for both the Source/Destination Interval for SheetB12
    Dim Isource As Integer
    Dim Idest As Integer
    Dim SourceInd As Integer
    Dim DestInd As Integer
    total_intervals = Sheets("S4").Range("H14").Value
    For Isource = 1 To total_intervals
        If U6a_PC_Available_Combobox.Column(0) = Worksheets("B7").Cells(7 + Isource, 2).Value And U6a_PC_Available_Combobox.Column(1) = Worksheets("B7").Cells(7 + Isource, 3).Value Then
        SourceInd = Isource + 7
        End If
    Next Isource
    For Idest = 1 To total_intervals
        If U6a_PC_Available_Combobox.Column(2) = Worksheets("B7").Cells(6, 3 + Idest).Value And U6a_PC_Available_Combobox.Column(3) = Worksheets("B7").Cells(7, 3 + Idest).Value Then
        DestInd = Idest + 3
        End If
    Next Idest

' [3] Assign a 1 in the Connectivity Table in Sheet B12
    Worksheets("B12").Cells(SourceInd, DestInd).Value = 1
    
' [4] Re-Draw Current Process Network
    Network_Drawing_Pathway_Config
    Range("A1").Select
    
' [5] Create a Custom Array to display for Current Pathway Combobox
    'Count the number of chosen connections
    ConCount = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(8, 4), Worksheets("B12").Cells(8 + total_intervals, 4 + total_intervals)))
    Dim PC_Chosen()
    Dim k As Integer
    ReDim PC_Chosen(ConCount, 4)
    
    k = 1
    For i = 1 To total_intervals
        For j = 1 To total_intervals
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                PC_Chosen(k - 1, 0) = Worksheets("B12").Cells(7 + i, 2).Value
                PC_Chosen(k - 1, 1) = Worksheets("B12").Cells(7 + i, 3).Value
                PC_Chosen(k - 1, 2) = Worksheets("B12").Cells(6, 3 + j).Value
                PC_Chosen(k - 1, 3) = Worksheets("B12").Cells(7, 3 + j).Value
                k = k + 1
            End If
        Next j
    Next i
    Me.U6a_PCSelected_Combobox.list = PC_Chosen

' [6] Export Number of Connections
    Worksheets("B12").Range("H2").Value = ConCount
End Sub
Private Sub U6a_SC_Add_Click()
On Error Resume Next
' [1] Make sure a connection has been selected
    If Me.U6a_SC_Available_Combobox.Value = "" Then
        MsgBox "Please select a Secondary Stream Connection to add to current pathway!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Get the Indexes for both the Source/Destination Interval for SheetB12
    Dim Isource As Integer
    Dim Idest As Integer
    Dim SourceInd As Integer
    Dim DestInd As Integer
    total_intervals = Sheets("S4").Range("H14").Value
    For Isource = 1 To total_intervals
        If U6a_SC_Available_Combobox.Column(0) = Worksheets("B7").Cells(12 + total_intervals + Isource, 2).Value And U6a_SC_Available_Combobox.Column(1) = Worksheets("B7").Cells(12 + total_intervals + Isource, 3).Value Then
        SourceInd = Isource + 12 + total_intervals
        End If
    Next Isource
    For Idest = 1 To total_intervals
        If U6a_SC_Available_Combobox.Column(2) = Worksheets("B7").Cells(11 + total_intervals, 3 + Idest).Value And U6a_SC_Available_Combobox.Column(3) = Worksheets("B7").Cells(12 + total_intervals, 3 + Idest).Value Then
        DestInd = Idest + 3
        End If
    Next Idest

' [3] Assign a 1 in the Connectivity Table in Sheet B12
    Worksheets("B12").Cells(SourceInd, DestInd).Value = 1
    
' [4] Re-Draw Current Process Network
    Network_Drawing_Pathway_Config
    Range("A1").Select
    
' [5] Create a Custom Array to display for Current Pathway Combobox
    'Count the number of chosen connections
    ConCount = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(13 + total_intervals, 4), Worksheets("B12").Cells(13 + total_intervals + total_intervals, 4 + total_intervals)))
    Dim SC_Chosen()
    Dim k As Integer
    ReDim SC_Chosen(ConCount, 4)
    
    k = 1
    For i = 1 To total_intervals
        For j = 1 To total_intervals
            If Worksheets("B12").Cells(12 + total_intervals + i, 3 + j).Value = 1 Then
                SC_Chosen(k - 1, 0) = Worksheets("B12").Cells(12 + total_intervals + i, 2).Value
                SC_Chosen(k - 1, 1) = Worksheets("B12").Cells(12 + total_intervals + i, 3).Value
                SC_Chosen(k - 1, 2) = Worksheets("B12").Cells(11 + total_intervals, 3 + j).Value
                SC_Chosen(k - 1, 3) = Worksheets("B12").Cells(12 + total_intervals, 3 + j).Value
                k = k + 1
            End If
        Next j
    Next i
    Me.U6a_SCSelected_Combobox.list = SC_Chosen

' [6] Export Number of Connections
    Worksheets("B12").Range("J2").Value = ConCount
End Sub






'***DISPLAY AVAILABLE CONNECTIONS***
Private Sub U6a_PC_Available_Combobox_Change()
On Error Resume Next
    If Me.U6a_PC_Available_Combobox.Value = "" Then
        MsgBox "Please select a valid Primary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
    If Val(U6a_PC_Available_Combobox.Column(0)) = 1 Then
        U6a_PC_Display.Text = "Source FEED" & "-" & U6a_PC_Available_Combobox.Column(1) & "    to    " & "Destination " & U6a_PC_Available_Combobox.Column(2) & "-" & U6a_PC_Available_Combobox.Column(3)
    ElseIf Val(U6a_PC_Available_Combobox.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U6a_PC_Display.Text = "Source " & U6a_PC_Available_Combobox.Column(0) & "-" & U6a_PC_Available_Combobox.Column(1) & "    to    " & "Destination PROD" & "-" & U6a_PC_Available_Combobox.Column(3)
    Else
    U6a_PC_Display.Text = "Source " & U6a_PC_Available_Combobox.Column(0) & "-" & U6a_PC_Available_Combobox.Column(1) & "    to    " & "Destination " & U6a_PC_Available_Combobox.Column(2) & "-" & U6a_PC_Available_Combobox.Column(3)
    End If
End Sub
Private Sub U6a_SC_Available_Combobox_Change()
On Error Resume Next
    If Me.U6a_SC_Available_Combobox.Value = "" Then
        MsgBox "Please select a valid Secondary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
    If Val(U6a_SC_Available_Combobox.Column(0)) = 1 Then
        U6a_SC_Display.Text = "Source FEED" & "-" & U6a_SC_Available_Combobox.Column(1) & "    to    " & "Destination " & U6a_SC_Available_Combobox.Column(2) & "-" & U6a_SC_Available_Combobox.Column(3)
    ElseIf Val(U6a_SC_Available_Combobox.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U6a_SC_Display.Text = "Source " & U6a_SC_Available_Combobox.Column(0) & "-" & U6a_SC_Available_Combobox.Column(1) & "    to    " & "Destination PROD" & "-" & U6a_SC_Available_Combobox.Column(3)
    Else
    U6a_SC_Display.Text = "Source " & U6a_SC_Available_Combobox.Column(0) & "-" & U6a_SC_Available_Combobox.Column(1) & "    to    " & "Destination " & U6a_SC_Available_Combobox.Column(2) & "-" & U6a_SC_Available_Combobox.Column(3)
    End If
End Sub





'***CHOOSE FROM SELECTED PROCESS STREAMS***
Private Sub U6a_PCSelected_Combobox_Change()
On Error Resume Next
    If Me.U6a_PCSelected_Combobox.Value = "" Then
        MsgBox "Please select a valid Primary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If

    If Val(U6a_PCSelected_Combobox.Column(0)) = 1 Then
        U6a_PCSelected_Display = "Source FEED" & "-" & U6a_PCSelected_Combobox.Column(1) & "    to    " & "Destination " & U6a_PCSelected_Combobox.Column(2) & "-" & U6a_PCSelected_Combobox.Column(3)
    ElseIf Val(U6a_PCSelected_Combobox.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U6a_PCSelected_Display = "Source " & U6a_PCSelected_Combobox.Column(0) & "-" & U6a_PCSelected_Combobox.Column(1) & "    to    " & "Destination PROD" & "-" & U6a_PCSelected_Combobox.Column(3)
    Else
    U6a_PCSelected_Display = "Source " & U6a_PCSelected_Combobox.Column(0) & "-" & U6a_PCSelected_Combobox.Column(1) & "    to    " & "Destination " & U6a_PCSelected_Combobox.Column(2) & "-" & U6a_PCSelected_Combobox.Column(3)
    End If
End Sub
Private Sub U6a_SCSelected_Combobox_Change()
On Error Resume Next
    If Me.U6a_SCSelected_Combobox.Value = "" Then
        MsgBox "Please select a valid Secondary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If

    If Val(U6a_SCSelected_Combobox.Column(0)) = 1 Then
        U6a_SCSelected_Display = "Source FEED" & "-" & U6a_SCSelected_Combobox.Column(1) & "    to    " & "Destination " & U6a_SCSelected_Combobox.Column(2) & "-" & U6a_SCSelected_Combobox.Column(3)
    ElseIf Val(U6a_SCSelected_Combobox.Column(2)) = Worksheets("S4").Cells(12, 8).Value + 2 Then
        U6a_SCSelected_Display = "Source " & U6a_SCSelected_Combobox.Column(0) & "-" & U6a_SCSelected_Combobox.Column(1) & "    to    " & "Destination PROD" & "-" & U6a_SCSelected_Combobox.Column(3)
    Else
    U6a_SCSelected_Display = "Source " & U6a_SCSelected_Combobox.Column(0) & "-" & U6a_SCSelected_Combobox.Column(1) & "    to    " & "Destination " & U6a_SCSelected_Combobox.Column(2) & "-" & U6a_SCSelected_Combobox.Column(3)
    End If
End Sub





'***REMOVE CONNECTION FROM SELECTED***
Private Sub U6a_PC_Remove_Click()
On Error Resume Next
' [1] Make sure that a connection from Selected has been chosen
    If Me.U6a_PCSelected_Combobox.Value = "" Then
        MsgBox "Please select a valid Primary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Get the Indexes for both the Source/Destination Interval for SheetB12
    Dim Isource As Integer
    Dim Idest As Integer
    Dim SourceInd As Integer
    Dim DestInd As Integer
    total_intervals = Sheets("S4").Range("H14").Value
    For Isource = 1 To total_intervals
        If U6a_PCSelected_Combobox.Column(0) = Worksheets("B7").Cells(7 + Isource, 2).Value And U6a_PCSelected_Combobox.Column(1) = Worksheets("B7").Cells(7 + Isource, 3).Value Then
        SourceInd = Isource + 7
        End If
    Next Isource
    For Idest = 1 To total_intervals
        If U6a_PCSelected_Combobox.Column(2) = Worksheets("B7").Cells(6, 3 + Idest).Value And U6a_PCSelected_Combobox.Column(3) = Worksheets("B7").Cells(7, 3 + Idest).Value Then
        DestInd = Idest + 3
        End If
    Next Idest

' [3] Set the Cell Value to Empty (AKA, deleted)
    Worksheets("B12").Cells(SourceInd, DestInd).Value = ""
    
' [4] Re-Draw Current Process Network
    Network_Drawing_Pathway_Config
    Range("A1").Select

' [5] Update the Combobox Array
    ConCount = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(8, 4), Worksheets("B12").Cells(8 + total_intervals, 4 + total_intervals)))
    Dim PC_Chosen()
    Dim k As Integer
    ReDim PC_Chosen(ConCount, 4)
    
    k = 1
    For i = 1 To total_intervals
        For j = 1 To total_intervals
            If Worksheets("B12").Cells(7 + i, 3 + j).Value = 1 Then
                PC_Chosen(k - 1, 0) = Worksheets("B12").Cells(7 + i, 2).Value
                PC_Chosen(k - 1, 1) = Worksheets("B12").Cells(7 + i, 3).Value
                PC_Chosen(k - 1, 2) = Worksheets("B12").Cells(6, 3 + j).Value
                PC_Chosen(k - 1, 3) = Worksheets("B12").Cells(7, 3 + j).Value
                k = k + 1
            End If
        Next j
    Next i
    Me.U6a_PCSelected_Combobox.list = PC_Chosen
    MsgBox "Primary Stream Connection Deleted", , "TIPEM- Notice"

' [6] Export Number of Connections
    Worksheets("B12").Range("H2").Value = ConCount
End Sub
Private Sub U6a_SC_Remove_Click()
On Error Resume Next
' [1] Make sure that a connection from Selected has been chosen
    If Me.U6a_SCSelected_Combobox.Value = "" Then
        MsgBox "Please select a valid Secondary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Get the Indexes for both the Source/Destination Interval for SheetB12
    Dim Isource As Integer
    Dim Idest As Integer
    Dim SourceInd As Integer
    Dim DestInd As Integer
    total_intervals = Sheets("S4").Range("H14").Value
    For Isource = 1 To total_intervals
        If U6a_SCSelected_Combobox.Column(0) = Worksheets("B7").Cells(12 + total_intervals + Isource, 2).Value And U6a_SCSelected_Combobox.Column(1) = Worksheets("B7").Cells(12 + total_intervals + Isource, 3).Value Then
        SourceInd = Isource + 12 + total_intervals
        End If
    Next Isource
    For Idest = 1 To total_intervals
        If U6a_SCSelected_Combobox.Column(2) = Worksheets("B7").Cells(11 + total_intervals, 3 + Idest).Value And U6a_SCSelected_Combobox.Column(3) = Worksheets("B7").Cells(12 + total_intervals, 3 + Idest).Value Then
        DestInd = Idest + 3
        End If
    Next Idest

' [3] Set the Cell Value to Empty (AKA, Deleted)
    Worksheets("B12").Cells(SourceInd, DestInd).Value = ""
    
' [4] Re-Draw Current Process Network
    Network_Drawing_Pathway_Config
    Range("A1").Select
    
' [5] Create a Custom Array to display for Current Pathway Combobox
    'Count the number of chosen connections
    ConCount = Application.WorksheetFunction.CountA(Worksheets("B12").Range(Worksheets("B12").Cells(13 + total_intervals, 4), Worksheets("B12").Cells(13 + total_intervals + total_intervals, 4 + total_intervals)))
    Dim SC_Chosen()
    Dim k As Integer
    ReDim SC_Chosen(ConCount, 4)
    
    k = 1
    For i = 1 To total_intervals
        For j = 1 To total_intervals
            If Worksheets("B12").Cells(12 + total_intervals + i, 3 + j).Value = 1 Then
                SC_Chosen(k - 1, 0) = Worksheets("B12").Cells(12 + total_intervals + i, 2).Value
                SC_Chosen(k - 1, 1) = Worksheets("B12").Cells(12 + total_intervals + i, 3).Value
                SC_Chosen(k - 1, 2) = Worksheets("B12").Cells(11 + total_intervals, 3 + j).Value
                SC_Chosen(k - 1, 3) = Worksheets("B12").Cells(12 + total_intervals, 3 + j).Value
                k = k + 1
            End If
        Next j
    Next i
    Me.U6a_SCSelected_Combobox.list = SC_Chosen
    MsgBox "Secondary Stream Connection Deleted", , "TIPEM- Notice"

' [6] Export Number of Connections
    Worksheets("B12").Range("J2").Value = ConCount
End Sub





'***CLOSE USERFORM***
Private Sub U6a_Close1_Click()
Unload Me
End
End Sub
Private Sub U6a_Close2_Click()
Unload Me
End
End Sub
