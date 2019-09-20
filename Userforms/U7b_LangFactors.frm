VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} U7b_LangFactors 
   Caption         =   "Lang Factors for Capital and Operating Costs"
   ClientHeight    =   12840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8625
   OleObjectBlob   =   "U7b_LangFactors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "U7b_LangFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** FUNCTION TO UPDATE COST *****
Public Function U7b_UpdateCost()
    'Equipment Cost
    U7b_2EquipmentDelivery = Format(Worksheets("O4").Range("E7").Value, "$0,000.00")
    U7b_3TotalDelivered = Format(Worksheets("O4").Range("E8").Value, "$0,000.00")
    'Direct Capital Costs
    U7b_DC1 = Format(Worksheets("O4").Range("E10").Value, "$0,000.00")
    U7b_DC2 = Format(Worksheets("O4").Range("E11").Value, "$0,000.00")
    U7b_DC3 = Format(Worksheets("O4").Range("E12").Value, "$0,000.00")
    U7b_DC4 = Format(Worksheets("O4").Range("E13").Value, "$0,000.00")
    U7b_DC5 = Format(Worksheets("O4").Range("E14").Value, "$0,000.00")
    U7b_DC6 = Format(Worksheets("O4").Range("E15").Value, "$0,000.00")
    U7b_DC7 = Format(Worksheets("O4").Range("E16").Value, "$0,000.00")
    U7b_DCTotal = Format(Worksheets("O4").Range("E17").Value, "$0,000.00")
    'Indirect Capital Costs
    U7b_IDC1 = Format(Worksheets("O4").Range("E19").Value, "$0,000.00")
    U7b_IDC2 = Format(Worksheets("O4").Range("E20").Value, "$0,000.00")
    U7b_IDC3 = Format(Worksheets("O4").Range("E21").Value, "$0,000.00")
    U7b_IDC4 = Format(Worksheets("O4").Range("E22").Value, "$0,000.00")
    U7b_IDC5 = Format(Worksheets("O4").Range("E23").Value, "$0,000.00")
    U7b_IDCTotal = Format(Worksheets("O4").Range("E24").Value, "$0,000.00")
    'Fixed Capital
    U7b_FCI = Format(Worksheets("O4").Range("E25").Value, "$0,000.00")
    U7b_WC = Format(Worksheets("O4").Range("E26").Value, "$0,000.00")
    U7b_TCI = Format(Worksheets("O4").Range("E27").Value, "$0,000.00")
    'Fixed Operating Costs
    U7b_FOC1 = Format(Worksheets("O4").Range("J7").Value, "$0,000.00")
    U7b_FOC2 = Format(Worksheets("O4").Range("J8").Value, "$0,000.00")
    U7b_FOC3 = Format(Worksheets("O4").Range("J9").Value, "$0,000.00")
    U7b_FOC4 = Format(Worksheets("O4").Range("J10").Value, "$0,000.00")
    U7b_FOC5 = Format(Worksheets("O4").Range("J11").Value, "$0,000.00")
    U7b_FOC6 = Format(Worksheets("O4").Range("J12").Value, "$0,000.00")
    U7b_FOC7 = Format(Worksheets("O4").Range("J13").Value, "$0,000.00")
    U7b_FOC8 = Format(Worksheets("O4").Range("J14").Value, "$0,000.00")
    U7b_FOC_Total = Format(Worksheets("O4").Range("J15").Value, "$0,000.00")
End Function





'**** LOAD DEFAULT LANG FACTORS *****
Private Sub U7b_Default1_Click()
' [1] Enter Default Values into Userform
    'Direct Capital Costs
    Worksheets("O4").Range("C7").Value = 0.1            'Equipment Delivery Cost (EDC)
    Worksheets("O4").Range("C10").Value = 0.31          'Pipes, Fittings, Valves and Tanks
    Worksheets("O4").Range("C11").Value = 0.26          'Instrumentation and Controls
    Worksheets("O4").Range("C12").Value = 0.1           'Electrical Systems
    Worksheets("O4").Range("C13").Value = 0.05          'Machinery and Equipment
    Worksheets("O4").Range("C14").Value = 0.25          'Buildings
    Worksheets("O4").Range("C15").Value = 0.12          'Yard Improvement
    Worksheets("O4").Range("C16").Value = 0.5           'Service Facilities
    'Indirect Capital Costs
    Worksheets("O4").Range("C19").Value = 0.32          'Engineering and Supervision
    Worksheets("O4").Range("C20").Value = 0.39          'Construction and Installation
    Worksheets("O4").Range("C21").Value = 0.04          'Legal Expenses
    Worksheets("O4").Range("C22").Value = 0.19          'Contractor's Fees
    Worksheets("O4").Range("C23").Value = 0.3           'Contingency and Insurance
    'Working Capital
    Worksheets("O4").Range("C26").Value = 0.1           'Working Capital

' [2] Load Values into Userform
    Dim TPEC As Double
    TPEC = Worksheets("O4").Range("E6").Value
    'Direct Capital Costs
    U7b_EDC = Worksheets("O4").Range("C7").Value
    U7b_Pipings = Worksheets("O4").Range("C10").Value
    U7b_Instrumentation = Worksheets("O4").Range("C11").Value
    U7b_Electrical = Worksheets("O4").Range("C12").Value
    U7b_Machinery = Worksheets("O4").Range("C13").Value
    U7b_Buildings = Worksheets("O4").Range("C14").Value
    U7b_Yard = Worksheets("O4").Range("C15").Value
    U7b_Service = Worksheets("O4").Range("C16").Value
    'Indirect Capital Costs
    U7b_Engineering = Worksheets("O4").Range("C19").Value
    U7b_Construction = Worksheets("O4").Range("C20").Value
    U7b_Legal = Worksheets("O4").Range("C21").Value
    U7b_Contractor = Worksheets("O4").Range("C22").Value
    U7b_Contingency = Worksheets("O4").Range("C23").Value
    'Working Capital
    U7b_WorkingCapital = Worksheets("O4").Range("C26").Value

' [3] Calculate Cost Factors
    'Equipment Cost
    U7b_2EquipmentDelivery = Format(Worksheets("O4").Range("E7").Value, "$0,000.00")
    U7b_3TotalDelivered = Format(Worksheets("O4").Range("E8").Value, "$0,000.00")
    'Direct Capital Costs
    U7b_DC1 = Format(Worksheets("O4").Range("E10").Value, "$0,000.00")
    U7b_DC2 = Format(Worksheets("O4").Range("E11").Value, "$0,000.00")
    U7b_DC3 = Format(Worksheets("O4").Range("E12").Value, "$0,000.00")
    U7b_DC4 = Format(Worksheets("O4").Range("E13").Value, "$0,000.00")
    U7b_DC5 = Format(Worksheets("O4").Range("E14").Value, "$0,000.00")
    U7b_DC6 = Format(Worksheets("O4").Range("E15").Value, "$0,000.00")
    U7b_DC7 = Format(Worksheets("O4").Range("E16").Value, "$0,000.00")
    U7b_DCTotal = Format(Worksheets("O4").Range("E17").Value, "$0,000.00")
    'Indirect Capital Costs
    U7b_IDC1 = Format(Worksheets("O4").Range("E19").Value, "$0,000.00")
    U7b_IDC2 = Format(Worksheets("O4").Range("E20").Value, "$0,000.00")
    U7b_IDC3 = Format(Worksheets("O4").Range("E21").Value, "$0,000.00")
    U7b_IDC4 = Format(Worksheets("O4").Range("E22").Value, "$0,000.00")
    U7b_IDC5 = Format(Worksheets("O4").Range("E23").Value, "$0,000.00")
    U7b_IDCTotal = Format(Worksheets("O4").Range("E24").Value, "$0,000.00")
    'Fixed Capital
    U7b_FCI = Format(Worksheets("O4").Range("E25").Value, "$0,000.00")
    U7b_WC = Format(Worksheets("O4").Range("E26").Value, "$0,000.00")
    U7b_TCI = Format(Worksheets("O4").Range("E27").Value, "$0,000.00")
End Sub
Private Sub U7b_Default2_Click()
' [1] Enter Default Values into Userform
    'Fixed Operating Costs
    Worksheets("O4").Range("H7").Value = 0.02            'Maintenance and Repairs
    Worksheets("O4").Range("H8").Value = 0.1             'Operating Labor
    Worksheets("O4").Range("H9").Value = 0.2             'Supervision
    Worksheets("O4").Range("H10").Value = 0.005           'Property Taxes
    Worksheets("O4").Range("H11").Value = 0.01           'Insurance
    Worksheets("O4").Range("H12").Value = 0.6            'Plant Overhead
    Worksheets("O4").Range("H13").Value = 0.2            'Laboratory
    Worksheets("O4").Range("H14").Value = 0.01           'Royalties
    
' [2] Load Values into Userform
    'Fixed Operating Costs
    U7b_MaR = Worksheets("O4").Range("H7").Value
    U7b_Labor = Worksheets("O4").Range("H8").Value
    U7b_Supervision = Worksheets("O4").Range("H9").Value
    U7b_Taxes = Worksheets("O4").Range("H10").Value
    U7b_Insurance = Worksheets("O4").Range("H11").Value
    U7b_Overhead = Worksheets("O4").Range("H12").Value
    U7b_Laboratory = Worksheets("O4").Range("H13").Value
    U7b_Royalties = Worksheets("O4").Range("H14").Value
    
' [3] Calculate Cost Factors
    'Fixed Operating Costs
    U7b_FOC1 = Format(Worksheets("O4").Range("J7").Value, "$0,000.00")
    U7b_FOC2 = Format(Worksheets("O4").Range("J8").Value, "$0,000.00")
    U7b_FOC3 = Format(Worksheets("O4").Range("J9").Value, "$0,000.00")
    U7b_FOC4 = Format(Worksheets("O4").Range("J10").Value, "$0,000.00")
    U7b_FOC5 = Format(Worksheets("O4").Range("J11").Value, "$0,000.00")
    U7b_FOC6 = Format(Worksheets("O4").Range("J12").Value, "$0,000.00")
    U7b_FOC7 = Format(Worksheets("O4").Range("J13").Value, "$0,000.00")
    U7b_FOC8 = Format(Worksheets("O4").Range("J14").Value, "$0,000.00")
    U7b_FOC_Total = Format(Worksheets("O4").Range("J15").Value, "$0,000.00")
End Sub





'**** CHANGING THE FACTOR VALUES ****
Private Sub U7b_EDC_Change()
    Worksheets("O4").Range("C7").Value = U7b_EDC.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Pipings_Change()
    Worksheets("O4").Range("C10").Value = U7b_Pipings.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Instrumentation_Change()
    Worksheets("O4").Range("C11").Value = U7b_Instrumentation.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Electrical_Change()
    Worksheets("O4").Range("C12").Value = U7b_Electrical.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Machinery_Change()
    Worksheets("O4").Range("C13").Value = U7b_Machinery.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Buildings_Change()
    Worksheets("O4").Range("C14").Value = U7b_Buildings.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Yard_Change()
    Worksheets("O4").Range("C15").Value = U7b_Yard.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Service_Change()
    Worksheets("O4").Range("C16").Value = U7b_Service.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Engineering_Change()
    Worksheets("O4").Range("C19").Value = U7b_Engineering.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Construction_Change()
    Worksheets("O4").Range("C20").Value = U7b_Construction.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Legal_Change()
    Worksheets("O4").Range("C21").Value = U7b_Legal.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Contractor_Change()
    Worksheets("O4").Range("C22").Value = U7b_Contractor.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Contingency_Change()
    Worksheets("O4").Range("C23").Value = U7b_Contingency.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_WorkingCapital_Change()
    Worksheets("O4").Range("C26").Value = U7b_WorkingCapital.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_MaR_Change()
    Worksheets("O4").Range("H7").Value = U7b_MaR.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Labor_Change()
    Worksheets("O4").Range("H8").Value = U7b_Labor.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Supervision_Change()
    Worksheets("O4").Range("H9").Value = U7b_Supervision.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Taxes_Change()
    Worksheets("O4").Range("H10").Value = U7b_Taxes.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Insurance_Change()
    Worksheets("O4").Range("H11").Value = U7b_Insurance.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Overhead_Change()
    Worksheets("O4").Range("H12").Value = U7b_Overhead.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Laboratory_Change()
    Worksheets("O4").Range("H13").Value = U7b_Laboratory.Value
    U7b_UpdateCost
End Sub
Private Sub U7b_Royalties_Change()
    Worksheets("O4").Range("H14").Value = U7b_Royalties.Value
    U7b_UpdateCost
End Sub





'**** UPON INITIALIZING, LOAD TPEC AND LOAD DEFAULT FACTOR VALUES ****
Private Sub UserForm_Initialize()
' [1] Declare Variables
    Dim steps As Integer
    Dim total_int As Integer
    Dim feed_int As Integer
    Dim product_int As Integer
    Dim process_int As Integer
    Dim TPEC As Double
    steps = Worksheets("S4").Range("H12").Value + 2
    total_int = Worksheets("S4").Range("H14").Value
    feed_int = Worksheets("S4").Range("F13").Value
    product_int = Worksheets("S4").Range("F" & 12 + steps).Value
    process_int = total_int - feed_int - product_int

' [2] Calculate Total Purchased Equipment Cost
    'Calculate Total Equipment Cost
    If Worksheets("O3").Range("F2").Value = 0 Then
        MsgBox "Equipment Costs for each Interval Need to be Specified!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    ElseIf Worksheets("O2").Range("F2").Value = 0 Then
        MsgBox "Mass Balances have not been calculated!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    TPEC = 0
    For i = 1 To process_int
        TPEC = TPEC + Worksheets("O3").Cells(6 + i, 8).Value
    Next i
    
    'Save TPEC Value to Userform
    Worksheets("O4").Range("E6").Value = TPEC
    
    'Load TPEC Value to Userform
    U7b_1EquipmentCost.Caption = Format(Worksheets("O4").Range("E6").Value, "$0,000.00")

' [3] Load Other Values into Userform
    'Direct Capital Costs
    U7b_EDC = Worksheets("O4").Range("C7").Value
    U7b_Pipings = Worksheets("O4").Range("C10").Value
    U7b_Instrumentation = Worksheets("O4").Range("C11").Value
    U7b_Electrical = Worksheets("O4").Range("C12").Value
    U7b_Machinery = Worksheets("O4").Range("C13").Value
    U7b_Buildings = Worksheets("O4").Range("C14").Value
    U7b_Yard = Worksheets("O4").Range("C15").Value
    U7b_Service = Worksheets("O4").Range("C16").Value
    'Indirect Capital Costs
    U7b_Engineering = Worksheets("O4").Range("C19").Value
    U7b_Construction = Worksheets("O4").Range("C20").Value
    U7b_Legal = Worksheets("O4").Range("C21").Value
    U7b_Contractor = Worksheets("O4").Range("C22").Value
    U7b_Contingency = Worksheets("O4").Range("C23").Value
    'Working Capital
    U7b_WorkingCapital = Worksheets("O4").Range("C26").Value
    'Fixed Operating Costs
    U7b_MaR = Worksheets("O4").Range("H7").Value
    U7b_Labor = Worksheets("O4").Range("H8").Value
    U7b_Supervision = Worksheets("O4").Range("H9").Value
    U7b_Taxes = Worksheets("O4").Range("H10").Value
    U7b_Insurance = Worksheets("O4").Range("H11").Value
    U7b_Overhead = Worksheets("O4").Range("H12").Value
    U7b_Laboratory = Worksheets("O4").Range("H13").Value
    U7b_Royalties = Worksheets("O4").Range("H14").Value

' [3] Calculate Cost Factors
    'Equipment Cost
    U7b_2EquipmentDelivery = Format(Worksheets("O4").Range("E7").Value, "$0,000.00")
    U7b_3TotalDelivered = Format(Worksheets("O4").Range("E8").Value, "$0,000.00")
    'Direct Capital Costs
    U7b_DC1 = Format(Worksheets("O4").Range("E10").Value, "$0,000.00")
    U7b_DC2 = Format(Worksheets("O4").Range("E11").Value, "$0,000.00")
    U7b_DC3 = Format(Worksheets("O4").Range("E12").Value, "$0,000.00")
    U7b_DC4 = Format(Worksheets("O4").Range("E13").Value, "$0,000.00")
    U7b_DC5 = Format(Worksheets("O4").Range("E14").Value, "$0,000.00")
    U7b_DC6 = Format(Worksheets("O4").Range("E15").Value, "$0,000.00")
    U7b_DC7 = Format(Worksheets("O4").Range("E16").Value, "$0,000.00")
    U7b_DCTotal = Format(Worksheets("O4").Range("E17").Value, "$0,000.00")
    'Indirect Capital Costs
    U7b_IDC1 = Format(Worksheets("O4").Range("E19").Value, "$0,000.00")
    U7b_IDC2 = Format(Worksheets("O4").Range("E20").Value, "$0,000.00")
    U7b_IDC3 = Format(Worksheets("O4").Range("E21").Value, "$0,000.00")
    U7b_IDC4 = Format(Worksheets("O4").Range("E22").Value, "$0,000.00")
    U7b_IDC5 = Format(Worksheets("O4").Range("E23").Value, "$0,000.00")
    U7b_IDCTotal = Format(Worksheets("O4").Range("E24").Value, "$0,000.00")
    'Fixed Capital
    U7b_FCI = Format(Worksheets("O4").Range("E25").Value, "$0,000.00")
    U7b_WC = Format(Worksheets("O4").Range("E26").Value, "$0,000.00")
    U7b_TCI = Format(Worksheets("O4").Range("E27").Value, "$0,000.00")
    'Fixed Operating Costs
    U7b_FOC1 = Format(Worksheets("O4").Range("J7").Value, "$0,000.00")
    U7b_FOC2 = Format(Worksheets("O4").Range("J8").Value, "$0,000.00")
    U7b_FOC3 = Format(Worksheets("O4").Range("J9").Value, "$0,000.00")
    U7b_FOC4 = Format(Worksheets("O4").Range("J10").Value, "$0,000.00")
    U7b_FOC5 = Format(Worksheets("O4").Range("J11").Value, "$0,000.00")
    U7b_FOC6 = Format(Worksheets("O4").Range("J12").Value, "$0,000.00")
    U7b_FOC7 = Format(Worksheets("O4").Range("J13").Value, "$0,000.00")
    U7b_FOC8 = Format(Worksheets("O4").Range("J14").Value, "$0,000.00")
    U7b_FOC_Total = Format(Worksheets("O4").Range("J15").Value, "$0,000.00")
End Sub





'**** CLOSE USERFORM *****
Private Sub U7b_Save_Click()
If Worksheets("O4").Range("E27").Value = "" Then
    MsgBox "Capital Cost has not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
    Exit Sub
Else
    Worksheets("O4").Range("F2").Value = 1
    MsgBox "Capital Cost Information has been Saved", vbExclamation, "TIPEM- Notice"
End If

Unload Me
End
End Sub
Private Sub U7b_Save2_Click()
If Worksheets("O4").Range("J15").Value = "" Or Worksheets("O4").Range("J22").Value = "" Then
    MsgBox "Operating Costs have not been Calculated Yet!!!", vbExclamation, "TIPEM- Error"
    Exit Sub
Else
    Worksheets("O4").Range("H2").Value = 1
    MsgBox "Operating Cost Information has been Saved", vbExclamation, "TIPEM- Notice"
End If

Unload Me
End
End Sub
