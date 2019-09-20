VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U5a_StreamsIn 
   Caption         =   "View Incoming Streams"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   OleObjectBlob   =   "U5a_StreamsIn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U5a_StreamsIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
    U5a_CurrentInt = "[" & Current_Step & "-" & Current_Int & "]   " & Worksheets("B10").Cells(Current_Row, 4)
    U5a_StreamsIn.Caption = "View Incoming Streams into " & "[" & Current_Step & "-" & Current_Int & "] " & Worksheets("B10").Cells(Current_Row, 4)


'[2] Find Current Interval In the Connectivity Matrix
    Dim Found_Column As Integer
    For b = 1 To Num_Int
        If Worksheets("B7").Cells(6, 3 + b).Value = Current_Step And Worksheets("B7").Cells(7, 3 + b).Value = Current_Int Then
            Found_Column = 3 + b
        End If
    Next b
    
    
'[3] If an Inteval is Connected to current Interval from Connectivity Matrix, append to ListBox
    U5a_PC_List.Clear
    U5a_SC_List.Clear
    U5a_PC_List.ColumnCount = 3
    U5a_SC_List.ColumnCount = 3
    
    Dim Start_Index_PC As Integer
    Dim Start_Index_SC As Integer
    Start_Index_PC = 0
    Start_Index_SC = 0

    For i = 1 To Num_Int
        If Worksheets("B7").Cells(7 + i, Found_Column).Value = 1 Then
            U5a_PC_List.AddItem
            U5a_PC_List.list(Start_Index_PC, 0) = Worksheets("B10").Cells(7 + i, 2).Value
            U5a_PC_List.list(Start_Index_PC, 1) = Worksheets("B10").Cells(7 + i, 3).Value
            U5a_PC_List.list(Start_Index_PC, 2) = Worksheets("B10").Cells(7 + i, 4).Value
            Start_Index_PC = Start_Index_PC + 1
        End If
    Next i
    
    For j = 1 To Num_Int
        If Worksheets("B7").Cells(12 + Num_Int + j, Found_Column).Value = 1 Then
            U5a_SC_List.AddItem
            U5a_SC_List.list(Start_Index_SC, 0) = Worksheets("B10").Cells(7 + j, 2).Value
            U5a_SC_List.list(Start_Index_SC, 1) = Worksheets("B10").Cells(7 + j, 3).Value
            U5a_SC_List.list(Start_Index_SC, 2) = Worksheets("B10").Cells(7 + j, 4).Value
            Start_Index_SC = Start_Index_SC + 1
        End If
    Next j
'***SIMULATE BUTTON CLICK***
Sheets("S5").Shapes.Range(Array("Oval 58")).ZOrder msoSendToBack
End Sub


Private Sub U5a_Button1_Click()
Unload Me
End
End Sub
