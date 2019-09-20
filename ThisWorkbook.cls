VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''   TIPEM INITIALIZATION MODULE   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ÀÌÁöÈ¯/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [INITIALIZE] WORKBOOK OPEN
Private Sub Workbook_Open()
    ' [GENERAL] Lock Workbook
    Worksheets("MAIN MENU").Protect Password:="lense0310"
    'Worksheets("S1").Protect Password:="lense0310"
    'Worksheets("S2").Protect Password:="lense0310"
    'Worksheets("S2").Protect Password:="lense0310"
    'Worksheets("S4").Protect Password:="lense0310"
    'Worksheets("S5").Protect Password:="lense0310"
    'Worksheets("S6").Protect Password:="lense0310"
    'Worksheets("S7").Protect Password:="lense0310"
    'Worksheets("S8").Protect Password:="lense0310"
    'Worksheets("MAIN MEN.Range("A2").Locked = True

    ' [GENERAL] LastSheet and ThisSheet for Macro to return users to last active sheet in main menu
    sLastSheet = ActiveSheet.Name
    sThisSheet = ActiveSheet.Name
    
    ' [GENERAL] Application Full Screen
    On Error Resume Next
    With Application
        .DisplayFullScreen = True
        .Width = 1400
        .Height = 900
    End With
    Application.ScreenUpdating = False
    Application.DisplayFormulaBar = False
    Application.CommandBars.ActiveMenuBar.Enabled = False
    Application.CommandBars("Full Screen").Enabled = False
    Application.OnKey "{ESC}", ""
    
    ' [GENERAL] Set Zoom on Full Screen
    Dim i
    Dim Current As Worksheet
    Set Current = ActiveSheet
    For i = 1 To ThisWorkbook.Sheets.Count
        Sheets(i).Activate
        ActiveWindow.Zoom = 110
        ' ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
        ' ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayFormulas = False
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayWhitespace = False
    Next
    
    ' [S0] Set MAIN MENU Sheet (Sheet 1) as Default upon Opening Program
    Worksheets("MAIN MENU").Activate
    
    ' [S1] Disable Scrollbar based on # of Materials
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
    
    ' [S2] Display Transportations Initial
    Set Transport_Data = Sheets("B5").Range("B5:D24")
    Set Transport_Display = Sheets("S2").Range("P15:R34")
    Transport_Display.Value = Transport_Data.Value
    
    ' [S4] Hide Overlay for Feedstock Specification
    Worksheets("S5").Shapes.Range(Array("Rectangle 3")).Visible = False
    Worksheets("S5").Shapes.Range(Array("Button 22")).Visible = False
    Worksheets("S5").Shapes.Range(Array("TextBox 69")).Visible = False
    Worksheets("S5").Shapes.Range(Array("TextBox 184")).ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Placeholder"
End Sub


' [INITIALIZE] WORKSHEET ACTIVATE
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    sLastSheet = sThisSheet
    sThisSheet = Sh.Name
End Sub

