Attribute VB_Name = "INFRA_Miscellaneous"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''    TIPEM MISCELLANEOUS MODULE   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ÀÌÁöÈ¯/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [PUBLIC VARIABLES]
Public sLastSheet As Variant
Public sThisSheet As Variant


' [INFRA] Return to Sheet
Public Sub I_ReturnToSheet()
    Sheets(sLastSheet).Activate
End Sub


' [INFRA] Save Project
Sub I_SaveAs()
    Dim TIPEMProjectName As Variant
    TIPEMProjectName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.XLSM), *.XLSM", Title:="Save TIPEM Project File")
    If TIPEMProjectName = False Then Exit Sub
    ActiveWorkbook.SaveAs FileName:=TIPEMProjectName
End Sub


' [INFRA] Activate Full Screen
Sub I_FullScreen()
    Application.DisplayFullScreen = True
End Sub


' [INFRA] Cancel Full Screen
Sub I_Exit_Fullscreen()
    Application.DisplayFullScreen = False
    ActiveWindow.Zoom = 110
End Sub


' [INFRA] Move to Next Stage of Project
' [1] Declare Variables
Sub I_GoSheetNext()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub


' [INFRA] Move to Next Stage for SYSTEM SPECIFICATION: Initialize Interval Specification Tables
' [1] Declare Variables
Sub I_GoSheetNext_SSpec()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Note before moving on to System Specification, make sure that the Process Network is finalized. If go back to
' process network (S3) then it will clear all existing Process Specification (S4) info!!
If MsgBox("Proceed with the current process network??", vbYesNo, "TIPEM- Warning") = vbNo Then Exit Sub
    
' [3] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
Range("A1").Select
End Sub


' [INFRA] Move to Previous Stage of Project
' [1] Declare Variables
Sub I_GoSheetBack()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet - lMove
        If lNext < 1 Then
            lMove = 0
            lNext = lSheets
            lSheet = lSheets
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub


' [INFRA] Move to Next Stage of Project (For Sheets S1)
Sub I_GoSheetNext_Materials()
' [1] Declare Variables
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Display Materials Data in Display Table
If Sheets("B2").Range("K3") >= 21 Then
    S1.ScrollBar2.Visible = True
    S1.ScrollBar2.Max = S3_2.UsedRange.Rows.Count - 19
    S1.ScrollBar2.Min = 4
    S1.ScrollBar2.Value = 5
    Else
        S1.ScrollBar2.Visible = False
        Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
        Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
        DisplayUpTo20.Value = MaterialsUpTo20.Value
End If

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub


' [INFRA] Move to Next Stage of Project (For Sheets S1)
Sub I_GoSheetNext_Materials2()
' [0] Make sure materials list is final
If MsgBox("Continue project with current materials list?? Adding or removing materials after this step may cause TIPEM to crash!!", vbYesNo, "TIPEM- Warning") = vbNo Then Exit Sub

' [1] Declare Variables
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Display Materials Data in Display Table
If Sheets("B2").Range("K3") >= 21 Then
    S1.ScrollBar2.Visible = True
    S1.ScrollBar2.Max = S3_2.UsedRange.Rows.Count - 19
    S1.ScrollBar2.Min = 4
    S1.ScrollBar2.Value = 5
    Else
        S1.ScrollBar2.Visible = False
        Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
        Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
        DisplayUpTo20.Value = MaterialsUpTo20.Value
End If

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub


' [INFRA] Move to Next Stage of Project (For Sheets S2)
Sub I_GoSheetNext_Utilities()
' [0] Make sure materials list is final
If MsgBox("Continue project with current utilities and transportation list?? Adding or removing items after this step may cause TIPEM to crash!!", vbYesNo, "TIPEM- Warning") = vbNo Then Exit Sub

' [1] Declare Variables
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub


' [INFRA] Move to Next Stage for LCA. **MAKE SURE TEA WAS CALCULATED**
' [1] Declare Variables
Sub I_GoSheetNext_LCA()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Note before moving on to System Specification, make sure that the Process Network is finalized. If go back to
' process network (S3) then it will clear all existing Process Specification (S4) info!!
If MsgBox("Make sure to ", vbYesNo, "TIPEM- Warning") = vbNo Then Exit Sub
    
' [3] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
Range("A1").Select
End Sub


' [INFRA] Move to Next Stage of Project (AFTER MASS BALANCES. FOR TEA AND LCA)
' [1] Declare Variables
Sub I_GoSheetNext_TEA()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] If CHECKSUM for MB, TEA, LCA = 1 and changes have been made to Process, Reset Tables
If MsgBox("Any changes made to the selected process will reset ALL results. Proceed???", vbYesNo, "TIPEM- Warning") = vbNo Then Exit Sub

' [3] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet + lMove
        If lNext > lSheets Then
            lMove = 0
            lNext = 1
            lSheet = 1
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
Range("A1").Select
End Sub


' [INFRA] Move to Previous Stage of Project (For Sheets S2, S3, S4)
' [1] Declare Variables
Sub I_GoSheetBack_Materials()
Dim wb As Workbook
Dim lSheets As Long
Dim lSheet As Long
Dim lMove As Long
Dim lNext As Long

' [2] Set Active Workbook
Set wb = ActiveWorkbook
lSheets = wb.Sheets.Count
lSheet = ActiveSheet.Index
lMove = 1

' [3] Display Materials Data in Display Table
If Sheets("B2").Range("K3") >= 21 Then
    S1.ScrollBar2.Visible = True
    S1.ScrollBar2.Max = S3_2.UsedRange.Rows.Count - 19
    S1.ScrollBar2.Min = 4
    S1.ScrollBar2.Value = 5
    Else
        S1.ScrollBar2.Visible = False
        Set MaterialsUpTo20 = Sheets("B2").Range("B4:I23")
        Set DisplayUpTo20 = Sheets("S1").Range("F13:M32")
        DisplayUpTo20.Value = MaterialsUpTo20.Value
End If

' [4] Move Worksheets
With wb
    For lMove = 1 To lSheets - 1
        lNext = lSheet - lMove
        If lNext < 1 Then
            lMove = 0
            lNext = lSheets
            lSheet = lSheets
        End If
        If .Sheets(lNext).Visible = True Then
            .Sheets(lNext).Select
            Exit For
        End If
    Next lMove
End With
End Sub
