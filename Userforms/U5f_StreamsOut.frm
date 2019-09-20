VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5f_StreamsOut 
   Caption         =   "Outgoing Streams from Interval #-#"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   OleObjectBlob   =   "U5f_StreamsOut.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5f_StreamsOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub U5f_Close_Click()
Unload Me
End
End Sub


Private Sub UserForm_Initialize()
'[1] Find current interval based on the indexes on top of Worksheet B10
    ' Declare Variables for Current Interval
    Dim Current_Step As Integer
    Dim Current_Int As Integer
    Dim Num_Int As Integer
    Dim Current_Row As Integer
    
    Current_Step = Worksheets("B10").Range("H3").Value
    Current_Int = Worksheets("B10").Range("K3").Value
    Num_Int = Worksheets("S4").Range("H14").Value
    
    ' Find the name of current interval from B10
    For a = 1 To Num_Int
        If Worksheets("B10").Cells(7 + a, 2).Value = Current_Step And Worksheets("B10").Cells(7 + a, 3).Value = Current_Int Then
            Current_Row = 7 + a
        End If
    Next a
    
    ' Display Name of Current Int in Textbox
    U5f_CurrentInt = "[" & Current_Step & "-" & Current_Int & "]   " & Worksheets("B10").Cells(Current_Row, 4)
    U5f_StreamsOut.Caption = "View Outgoing Streams from " & "[" & Current_Step & "-" & Current_Int & "] " & Worksheets("B10").Cells(Current_Row, 4)
    
    
'[2] If the Current Interval is Connected to an Outgoing Interval, Append to ListBox
    U5f_PC_List.Clear
    U5f_SC_List.Clear
    U5f_PC_List.ColumnCount = 3
    U5f_SC_List.ColumnCount = 3
    
    Dim Start_Index_PC As Integer
    Dim Start_Index_SC As Integer
    Start_Index_PC = 0
    Start_Index_SC = 0

    For i = 1 To Num_Int
        If Worksheets("B7").Cells(Current_Row, 3 + i).Value = 1 Then
            U5f_PC_List.AddItem
            U5f_PC_List.list(Start_Index_PC, 0) = Worksheets("B10").Cells(7 + i, 2).Value
            U5f_PC_List.list(Start_Index_PC, 1) = Worksheets("B10").Cells(7 + i, 3).Value
            U5f_PC_List.list(Start_Index_PC, 2) = Worksheets("B10").Cells(7 + i, 4).Value
            Start_Index_PC = Start_Index_PC + 1
        End If
    Next i
    
    For j = 1 To Num_Int
        If Worksheets("B7").Cells(Current_Row + 5 + Num_Int, 3 + j).Value = 1 Then
            U5f_SC_List.AddItem
            U5f_SC_List.list(Start_Index_SC, 0) = Worksheets("B10").Cells(7 + j, 2).Value
            U5f_SC_List.list(Start_Index_SC, 1) = Worksheets("B10").Cells(7 + j, 3).Value
            U5f_SC_List.list(Start_Index_SC, 2) = Worksheets("B10").Cells(7 + j, 4).Value
            Start_Index_SC = Start_Index_SC + 1
        End If
    Next j
'***SIMULATE BUTTON CLICK***
ActiveSheet.Shapes.Range(Array("Oval 66")).ZOrder msoSendToBack
ActiveSheet.Shapes.Range(Array("Oval 67")).ZOrder msoSendToBack
End Sub
