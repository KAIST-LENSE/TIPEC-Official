VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5j_TransportDistance 
   Caption         =   "Specify Transportation Distances between Intervals"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8505
   OleObjectBlob   =   "U5j_TransportDistance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5j_TransportDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub U5j_MultiPage1_Change()

End Sub

'***INITIALIZE USERFORM***
Private Sub UserForm_Initialize()
On Error Resume Next
' [1] Show list of available connection using B8 and B9 for Primary/Secondary Connections respectively
    ' [GENERAL VARIABLES]
    Dim num_PC As Integer
    Dim num_SC As Integer
    num_PC = Worksheets("B8").Range("C1").Value
    num_SC = Worksheets("B9").Range("C1").Value
    
    ' [PRIMARY CONNECTIONS] Custom Array
    Dim PC_array()
    ReDim PC_array(num_PC, 5)
    For a = 1 To num_PC
        PC_array(a - 1, 0) = Worksheets("B8").Cells(3 + a, 3).Value
        PC_array(a - 1, 1) = Worksheets("B8").Cells(3 + a, 4).Value
        PC_array(a - 1, 2) = Worksheets("B8").Cells(3 + a, 5).Value
        PC_array(a - 1, 3) = Worksheets("B8").Cells(3 + a, 6).Value
        PC_array(a - 1, 4) = 0
    Next a
    Me.U5j_PC_Transport_Combobox.list = PC_array

    ' [SECONDARY CONNECTIONS] Custom Array
    Dim SC_array()
    ReDim SC_array(num_SC, 5)
    For a = 1 To num_SC
        SC_array(a - 1, 0) = Worksheets("B9").Cells(3 + a, 3).Value
        SC_array(a - 1, 1) = Worksheets("B9").Cells(3 + a, 4).Value
        SC_array(a - 1, 2) = Worksheets("B9").Cells(3 + a, 5).Value
        SC_array(a - 1, 3) = Worksheets("B9").Cells(3 + a, 6).Value
        SC_array(a - 1, 4) = 0
    Next a
    Me.U5j_SC_Transport_Combobox.list = SC_array
End Sub




'***TRANSPORTATION METHOD SELECTED***
Private Sub U5j_PC_TransportList_Combobox_Change()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_PC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Display the selected transportation info in the textbox
    U5j_PC_TransportList_Display.Text = U5j_PC_TransportList_Combobox.Column(0) & "   |   " & U5j_PC_TransportList_Combobox.Column(1)

' [3] Assign Variables
    Transport_Index = Me.U5j_PC_TransportList_Combobox.Column(0)
    Num_Int = Worksheets("S4").Range("H14").Value
    Num_Spacing = 6
    Starting_Index = (Transport_Index * Num_Spacing) + ((Transport_Index - 1) * (Num_Int))
    
' [4] Update the Array showing List of PC and Specified Distances for that transportation
    Dim num_PC As Integer
    Dim PC_array()
    num_PC = Worksheets("B8").Range("C1").Value
    ReDim PC_array(num_PC, 5)
    For a = 1 To num_PC
        PC_array(a - 1, 0) = Worksheets("B8").Cells(3 + a, 3).Value
        PC_array(a - 1, 1) = Worksheets("B8").Cells(3 + a, 4).Value
        PC_array(a - 1, 2) = Worksheets("B8").Cells(3 + a, 5).Value
        PC_array(a - 1, 3) = Worksheets("B8").Cells(3 + a, 6).Value
        For i = 1 To Num_Int
            For j = 1 To Num_Int
                If Worksheets("B11").Cells(Starting_Index + i, 2).Value = PC_array(a - 1, 0) And Worksheets("B11").Cells(Starting_Index + i, 3).Value = PC_array(a - 1, 1) And Worksheets("B11").Cells(Starting_Index - 1, 3 + j).Value = PC_array(a - 1, 2) And Worksheets("B11").Cells(Starting_Index, 3 + j).Value = PC_array(a - 1, 3) Then
                    PC_array(a - 1, 4) = Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value
                    If Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value = "" Then
                        PC_array(a - 1, 4) = 0
                        Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value = 0
                    End If
                End If
            Next j
        Next i
    Next a
    Me.U5j_PC_Transport_Combobox.list = PC_array
End Sub
Private Sub U5j_SC_TransportList_Combobox_Change()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_SC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Display the selected transportation info in the textbox
    U5j_SC_TransportList_Display.Text = U5j_SC_TransportList_Combobox.Column(0) & "   |   " & U5j_SC_TransportList_Combobox.Column(1)

' [3] Assign Variables
    Transport_Index = Me.U5j_SC_TransportList_Combobox.Column(0)
    Num_Int = Worksheets("S4").Range("H14").Value
    Num_Spacing = 6
    Starting_Index = (Transport_Index * Num_Spacing) + ((Transport_Index - 1) * (Num_Int))
    
' [4] Update the Array showing List of PC and Specified Distances for that transportation
    Dim num_SC As Integer
    Dim SC_array()
    num_SC = Worksheets("B9").Range("C1").Value
    ReDim SC_array(num_SC, 5)
    For a = 1 To num_SC
        SC_array(a - 1, 0) = Worksheets("B9").Cells(3 + a, 3).Value
        SC_array(a - 1, 1) = Worksheets("B9").Cells(3 + a, 4).Value
        SC_array(a - 1, 2) = Worksheets("B9").Cells(3 + a, 5).Value
        SC_array(a - 1, 3) = Worksheets("B9").Cells(3 + a, 6).Value
        For i = 1 To Num_Int
            For j = 1 To Num_Int
                If Worksheets("B11").Cells(Starting_Index + i, 2).Value = SC_array(a - 1, 0) And Worksheets("B11").Cells(Starting_Index + i, 3).Value = SC_array(a - 1, 1) And Worksheets("B11").Cells(Starting_Index - 1, 3 + Num_Int + j).Value = SC_array(a - 1, 2) And Worksheets("B11").Cells(Starting_Index, 3 + Num_Int + j).Value = SC_array(a - 1, 3) Then
                    SC_array(a - 1, 4) = Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value
                    If Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value = "" Then
                        SC_array(a - 1, 4) = 0
                        Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value = 0
                    End If
                End If
            Next j
        Next i
    Next a
    Me.U5j_SC_Transport_Combobox.list = SC_array
End Sub




'***PRIMARY/SECONDARY STREAM CONNECTION SELECTED***
Private Sub U5j_PC_Transport_Combobox_Change()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_PC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U5j_PC_Transport_Combobox = "" Then
        MsgBox "No stream specified!! Make sure to select a Primary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Display Selected Primary Stream and its Transport Distance
    U5j_PC_Transport_Display = "Source Step [" & U5j_PC_Transport_Combobox.Column(0) & "-" & U5j_PC_Transport_Combobox.Column(1) & "]           to          Destination Step [" & U5j_PC_Transport_Combobox.Column(2) & "-" & U5j_PC_Transport_Combobox.Column(3) & "]                    " & U5j_PC_Transport_Combobox.Column(4) & "km"

' [3] Display Existing Transport Distance in TextBox
    U5j_PC_TransportDistance = U5j_PC_Transport_Combobox.Column(4)
End Sub
Private Sub U5j_SC_Transport_Combobox_Change()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_SC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U5j_SC_Transport_Combobox = "" Then
        MsgBox "No stream specified!! Make sure to select a Secondary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Display Selected Secondary Stream and its Transport Distance
    U5j_SC_Transport_Display = "Source Step [" & U5j_SC_Transport_Combobox.Column(0) & "-" & U5j_SC_Transport_Combobox.Column(1) & "]           to          Destination Step [" & U5j_SC_Transport_Combobox.Column(2) & "-" & U5j_SC_Transport_Combobox.Column(3) & "]                    " & U5j_SC_Transport_Combobox.Column(4) & "km"

' [3] Display Existing Transport Distance in TextBox
    U5j_SC_TransportDistance = U5j_SC_Transport_Combobox.Column(4)
End Sub




'***APPLY THE SPECIFIED TRANSPORTATION DISTANCE AND UPDATE THE DISPLAY ARRAY***
Private Sub U5j_PC_Apply_Click()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_PC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U5j_PC_Transport_Combobox = "" Then
        MsgBox "No stream specified!! Make sure to select a Primary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Make sure transportation distance is a Non-Negative Number
    If Me.U5j_PC_TransportDistance < 0 Then
        MsgBox "Transportation Distance cannot be a negative number!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [3] Assign Variables
    Transport_Index = Me.U5j_PC_TransportList_Combobox.Column(0)
    Num_Int = Worksheets("S4").Range("H14").Value
    Num_Spacing = 6
    Starting_Index = (Transport_Index * Num_Spacing) + ((Transport_Index - 1) * (Num_Int))

' [4] Save specified Transportation Distance in the correct cell
    For aa = 1 To Num_Int
        If Worksheets("B11").Cells(Starting_Index + aa, 2).Value = U5j_PC_Transport_Combobox.Column(0) And Worksheets("B11").Cells(Starting_Index + aa, 3).Value = U5j_PC_Transport_Combobox.Column(1) Then
            Current_Row = Starting_Index + aa
        End If
    Next aa
    For bb = 1 To Num_Int
        If Worksheets("B11").Cells(Starting_Index - 1, 3 + bb).Value = U5j_PC_Transport_Combobox.Column(2) And Worksheets("B11").Cells(Starting_Index, 3 + bb).Value = U5j_PC_Transport_Combobox.Column(3) Then
            Current_Col = 3 + bb
        End If
    Next bb
    Worksheets("B11").Cells(Current_Row, Current_Col).Value = U5j_PC_TransportDistance.Value

' [5] Update the Array showing List of PC and Specified Distances for that transportation
    Dim num_PC As Integer
    Dim PC_array()
    num_PC = Worksheets("B8").Range("C1").Value
    ReDim PC_array(num_PC, 5)
    For a = 1 To num_PC
        PC_array(a - 1, 0) = Worksheets("B8").Cells(3 + a, 3).Value
        PC_array(a - 1, 1) = Worksheets("B8").Cells(3 + a, 4).Value
        PC_array(a - 1, 2) = Worksheets("B8").Cells(3 + a, 5).Value
        PC_array(a - 1, 3) = Worksheets("B8").Cells(3 + a, 6).Value
        For i = 1 To Num_Int
            For j = 1 To Num_Int
                If Worksheets("B11").Cells(Starting_Index + i, 2).Value = PC_array(a - 1, 0) And Worksheets("B11").Cells(Starting_Index + i, 3).Value = PC_array(a - 1, 1) And Worksheets("B11").Cells(Starting_Index - 1, 3 + j).Value = PC_array(a - 1, 2) And Worksheets("B11").Cells(Starting_Index, 3 + j).Value = PC_array(a - 1, 3) Then
                    PC_array(a - 1, 4) = Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value
                    If Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value = "" Then
                        PC_array(a - 1, 4) = 0
                        Worksheets("B11").Cells(Starting_Index + i, 3 + j).Value = 0
                    End If
                End If
            Next j
        Next i
    Next a
    Me.U5j_PC_Transport_Combobox.list = PC_array
End Sub
Private Sub U5j_SC_Apply_Click()
On Error Resume Next
' [1] Make sure that a Transport Mode has first been selected
    If Me.U5j_SC_TransportList_Combobox = "" Then
        MsgBox "No Transport Mode specified!! Please choose a Mode of Transportation first!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    If Me.U5j_SC_Transport_Combobox = "" Then
        MsgBox "No stream specified!! Make sure to select a Secondary Stream Connection!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [2] Make sure transportation distance is a Non-Negative Number
    If Me.U5j_SC_TransportDistance < 0 Then
        MsgBox "Transportation Distance cannot be a negative number!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    
' [3] Assign Variables
    Transport_Index = Me.U5j_SC_TransportList_Combobox.Column(0)
    Num_Int = Worksheets("S4").Range("H14").Value
    Num_Spacing = 6
    Starting_Index = (Transport_Index * Num_Spacing) + ((Transport_Index - 1) * (Num_Int))

' [4] Save specified Transportation Distance in the correct cell
    For aa = 1 To Num_Int
        If Worksheets("B11").Cells(Starting_Index + aa, 2).Value = U5j_SC_Transport_Combobox.Column(0) And Worksheets("B11").Cells(Starting_Index + aa, 3).Value = U5j_SC_Transport_Combobox.Column(1) Then
            Current_Row = Starting_Index + aa
        End If
    Next aa
    For bb = 1 To Num_Int
        If Worksheets("B11").Cells(Starting_Index - 1, 3 + Num_Int + bb).Value = U5j_SC_Transport_Combobox.Column(2) And Worksheets("B11").Cells(Starting_Index, 3 + Num_Int + bb).Value = U5j_SC_Transport_Combobox.Column(3) Then
            Current_Col = 3 + Num_Int + bb
        End If
    Next bb
    Worksheets("B11").Cells(Current_Row, Current_Col).Value = U5j_SC_TransportDistance.Value

' [5] Update the Array showing List of SC and Specified Distances for that transportation
    Dim num_SC As Integer
    Dim SC_array()
    num_SC = Worksheets("B9").Range("C1").Value
    ReDim SC_array(num_SC, 5)
    For a = 1 To num_SC
        SC_array(a - 1, 0) = Worksheets("B9").Cells(3 + a, 3).Value
        SC_array(a - 1, 1) = Worksheets("B9").Cells(3 + a, 4).Value
        SC_array(a - 1, 2) = Worksheets("B9").Cells(3 + a, 5).Value
        SC_array(a - 1, 3) = Worksheets("B9").Cells(3 + a, 6).Value
        For i = 1 To Num_Int
            For j = 1 To Num_Int
                If Worksheets("B11").Cells(Starting_Index + i, 2).Value = SC_array(a - 1, 0) And Worksheets("B11").Cells(Starting_Index + i, 3).Value = SC_array(a - 1, 1) And Worksheets("B11").Cells(Starting_Index - 1, 3 + Num_Int + j).Value = SC_array(a - 1, 2) And Worksheets("B11").Cells(Starting_Index, 3 + Num_Int + j).Value = SC_array(a - 1, 3) Then
                    SC_array(a - 1, 4) = Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value
                    If Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value = "" Then
                        SC_array(a - 1, 4) = 0
                        Worksheets("B11").Cells(Starting_Index + i, 3 + Num_Int + j).Value = 0
                    End If
                End If
            Next j
        Next i
    Next a
    Me.U5j_SC_Transport_Combobox.list = SC_array
End Sub




'***CLOSE USERFORM***
Private Sub U5j_Close1_Click()
Unload Me
End
End Sub
Private Sub U5j_Close2_Click()
Unload Me
End
End Sub

