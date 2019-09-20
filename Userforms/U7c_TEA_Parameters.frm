VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U7c_TEA_Parameters 
   Caption         =   "Specify TEA Parameters"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   OleObjectBlob   =   "U7c_TEA_Parameters.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U7c_TEA_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** CLOSE USERFORM *****
Private Sub U7c_Close_Click()
' [1] Check if Material Balance, CAPEX, and Fixed OPEX have been Calculated
    If Worksheets("O4").Range("E27").Value = "" Then
        MsgBox "Capital Cost has not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O4").Range("J15").Value = "" Then
        MsgBox "Operating Cost has not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O4").Range("E6").Value = "" Then
        MsgBox "Equipment Costs need to be Specified!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If

' [2] Save Values to Userform
    Worksheets("O4").Range("H26").Value = U7c_Batch.Value
    Worksheets("O4").Range("H27").Value = U7c_Lifetime.Value
    Worksheets("O4").Range("H28").Value = U7c_Salvage.Value
    Worksheets("O4").Range("H29").Value = U7c_Tax.Value
    Worksheets("O4").Range("H30").Value = U7c_Interest.Value
    'Generic Expenses
    Worksheets("O4").Range("C40").Value = U7c_GE_Sales.Value
    Worksheets("O4").Range("C41").Value = U7c_GE_RD.Value
    Worksheets("O4").Range("C42").Value = U7c_GE_Admin.Value

' [3] Check if Any Input Boxes are Empty
    If U7c_Batch.Value = "" Or U7c_Lifetime.Value = "" Or U7c_Depreciable.Value = "" Or U7c_Salvage.Value = "" Or U7c_Tax.Value = "" Or U7c_Interest.Value = "" Or U7c_GE_Sales.Value = "" Or U7c_GE_RD.Value = "" Or U7c_GE_Admin.Value = "" Then
        MsgBox "Make sure to specify ALL fields!!!", vbExclamation, "TIPEM- Notice"
    Else
        MsgBox "Parameters Saved", vbExclamation, "TIPEM- Notice"
        Unload Me
        End
    End If
End Sub





'**** UPON INITIALIZING, LOAD TPEC AND LOAD DEFAULT FACTOR VALUES ****
Private Sub UserForm_Initialize()
' [1] Check if Material Balance, CAPEX, and Fixed OPEX have been calculated
    If Worksheets("O3").Range("F2").Value = 0 Then
        MsgBox "Equipment Costs for each Interval Need to be Specified!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O2").Range("F2").Value = 0 Then
        MsgBox "Mass Balances have not been calculated!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O4").Range("E6").Value = "" Then
        MsgBox "Please specify Capital Cost Lang Factors!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O4").Range("E27").Value = "" Then
        MsgBox "Capital Cost has not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O4").Range("J15").Value = "" Then
        MsgBox "Operating Cost has not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If

' [2] Load current values into Userform
    'General Parameters
    U7c_Batch.Value = Worksheets("O4").Range("H26").Value
    U7c_Lifetime.Value = Worksheets("O4").Range("H27").Value
    U7c_Depreciable.Value = Format(Worksheets("O4").Range("E46").Value, "$0,000.00")
    U7c_Salvage.Value = Worksheets("O4").Range("H28").Value
    U7c_Tax.Value = Worksheets("O4").Range("H29").Value
    U7c_Interest.Value = Worksheets("O4").Range("H30").Value
    'Generic Expenses
    U7c_GE_Sales.Value = Worksheets("O4").Range("C40").Value
    U7c_GE_RD.Value = Worksheets("O4").Range("C41").Value
    U7c_GE_Admin.Value = Worksheets("O4").Range("C42").Value
End Sub





'**** LOAD DEFAULT VALUES INTO USERFORM ****
Private Sub U7c_Default_Click()
' [1] Enter Default Values into O4
    'General Parameters
    Worksheets("O4").Range("H27").Value = 30            'Plant Lifetime (years)
    Worksheets("O4").Range("H28").Value = 0             'Salvage Value ($)
    Worksheets("O4").Range("H29").Value = 20            'Income Tax (%)
    Worksheets("O4").Range("H30").Value = 7             'Interest Rate (%)
    'Generic Expenses
    Worksheets("O4").Range("C40").Value = 0.03          'Sales Expense
    Worksheets("O4").Range("C41").Value = 0.05          'R&D
    Worksheets("O4").Range("C42").Value = 0.03          'Administration

' [2] Load Default Values Into Userform
    'General Parameters
    U7c_Batch.Value = Worksheets("O4").Range("H26").Value
    U7c_Lifetime.Value = Worksheets("O4").Range("H27").Value
    U7c_Depreciable.Value = Format(Worksheets("O4").Range("E46").Value, "$0,000.00")
    U7c_Salvage.Value = Format(Worksheets("O4").Range("H28").Value, "$0,000.00")
    U7c_Tax.Value = Worksheets("O4").Range("H29").Value
    U7c_Interest.Value = Worksheets("O4").Range("H30").Value
    'Generic Expenses
    U7c_GE_Sales.Value = Worksheets("O4").Range("C40").Value
    U7c_GE_RD.Value = Worksheets("O4").Range("C41").Value
    U7c_GE_Admin.Value = Worksheets("O4").Range("C42").Value
End Sub
