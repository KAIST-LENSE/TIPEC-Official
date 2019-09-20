Attribute VB_Name = "M_ProcessSpecification"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''   TIPEM PROCESS SPEC MODULE     ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ÀÌÁöÈ¯/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' [MASS BALANCE MODEL] Remove Duplicate Process Intervals
Function uniqueArr(ParamArray myArr() As Variant) As Variant()
On Error Resume Next
    Dim dict As Object
    Dim V As Variant, W As Variant
    Dim i As Long

Set dict = CreateObject("Scripting.Dictionary")
For Each V In myArr 'loop through each myArr
    For Each W In V 'loop through the contents of each myArr
        If Not dict.exists(W) Then dict.Add W, W
    Next W
Next V
uniqueArr = dict.keys
End Function







' [MASS BALANCE MODEL] Adopted from Kosan Roh
Public Sub TIPEM_Calculate_MB()
On Error GoTo message
'---------------------------------------------------------------------------------
'-----------------  [I] INITIALIZATION  ------------------------------------------
'---------------------------------------------------------------------------------
'[0] Check if all of the Process Intervals have been Specified
    If MsgBox("This assumes all of the Process Intervals have been specified. Are you want to procede??", vbYesNo, "TIPEM- Notice") = vbNo Then Exit Sub
    Dim Start As Date
    Start = Timer
    ' Make sure all values have been specified
    If Worksheets("S8").Range("F12").Value = "" Or Worksheets("S8").Range("G12").Value = "" Then
        MsgBox "Please specify all required inputs!!!", vbExclamation, "TIPEM- Error"
        Exit Sub
    End If
    ' Delete Existing Mass Balance
    Worksheets("O1").Range("B4:CZ220").ClearContents
    Worksheets("O1").Range("B4:CZ220").Font.Bold = False
    Worksheets("O1").Range("B4:CZ220").UnMerge
    Worksheets("O1").Range("B4:CZ220").Borders(xlInsideVertical).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Borders(xlInsideHorizontal).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Borders(xlEdgeTop).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Borders(xlEdgeBottom).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Borders(xlEdgeRight).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Borders(xlEdgeLeft).LineStyle = xlNone
    Worksheets("O1").Range("B4:CZ220").Interior.Pattern = xlNone
    Worksheets("O1").Range("B4:CZ220").Interior.TintAndShade = 0
    Worksheets("O1").Range("B4:CZ220").Interior.PatternTintAndShade = 0
    Worksheets("O2").Range("B4:CZ220").ClearContents
    Worksheets("O2").Range("B4:CZ220").Font.Bold = False
    Worksheets("O2").Range("B4:CZ220").UnMerge
    Worksheets("O2").Range("B4:CZ220").Borders(xlInsideVertical).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Borders(xlInsideHorizontal).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Borders(xlEdgeTop).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Borders(xlEdgeBottom).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Borders(xlEdgeRight).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Borders(xlEdgeLeft).LineStyle = xlNone
    Worksheets("O2").Range("B4:CZ220").Interior.Pattern = xlNone
    Worksheets("O2").Range("B4:CZ220").Interior.TintAndShade = 0
    Worksheets("O2").Range("B4:CZ220").Interior.PatternTintAndShade = 0


'[1] Compile an array of process intervals involved with the specified process
    Dim m As Integer
    Dim mm As Integer
    Dim nn As Integer
    Dim PC_Count As Integer
    Dim SC_Count As Integer
    Dim IntCount As Integer

    PC_Count = Worksheets("B12").Range("H2").Value
    SC_Count = Worksheets("B12").Range("J2").Value

    Dim Interval_Array_PC()
    Dim Interval_Array_SC()
    Dim MainArray()
    ReDim Interval_Array_PC(1 To PC_Count)
    If SC_Count <> 0 Then
        ReDim Interval_Array_SC(1 To SC_Count)
    End If
    
    m = 1
    For mm = 1 To Worksheets("S3").Range("H14").Value
        For nn = 1 To Worksheets("S3").Range("H14").Value
            If Worksheets("B12").Cells(7 + mm, 3 + nn).Value = 1 Then
                Interval_Array_PC(m) = Worksheets("B10").Cells(7 + nn, 4).Value
                m = m + 1
            End If
        Next nn
    Next mm
    m = 1
    For mm = 1 To Worksheets("S3").Range("H14").Value
        For nn = 1 To Worksheets("S3").Range("H14").Value
            If Worksheets("B12").Cells(12 + Worksheets("S3").Range("H14") + mm, 3 + nn).Value = 1 Then
                Interval_Array_SC(m) = Worksheets("B10").Cells(7 + nn, 4).Value
                m = m + 1
            End If
        Next nn
    Next mm
    If SC_Count <> 0 Then
        MainArray = uniqueArr(Interval_Array_PC, Interval_Array_SC)
    Else
        MainArray = uniqueArr(Interval_Array_PC)
    End If
    IntCount = UBound(MainArray) - LBound(MainArray) + 1



'---------------------------------------------------------------------------------
'-----------------  [II] DEFINE AND REDEFINE VARIABLES  --------------------------
'---------------------------------------------------------------------------------
'[2] Dummy Variables (Implicit)
    Dim a As Integer
    Dim b As Integer
    Dim bb As Integer
    Dim bbb As Integer
    Dim C As Integer
    Dim d As Integer
    Dim n As Integer
    Dim x As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim Count As Integer
    Dim cell_selected As Range


'[3] Process Indices
    Dim k As Integer 'Set of Intervals as Source (In-coming)
    Dim kk As Integer 'Set of Intervals as Destination (Out-going)
    Dim steps As Integer 'Set of Processing Steps
    Dim i As Integer 'Set of Materials
    Dim ute As Integer 'Set of ENERGY utilities
    Dim utm As Integer 'Set of MASS utilities
    Dim tp As Integer 'Set of Transportations
    steps = Worksheets("S3").Range("H12").Value + 2
    i = Worksheets("B2").Range("K3").Value
    ute = Worksheets("B3").Range("C1").Value
    utm = Worksheets("B4").Range("C1").Value
    k = Worksheets("S3").Range("H14").Value
    tp = Worksheets("B5").Range("C1").Value
    kk = k
    
    
'[4] Component Identification
    Dim Mat_Name() As String 'Name of compounds
    Dim EUtility_Name() As String 'Name of Energy Utilities
    Dim MUtility_Name() As String 'Name of Mass Utilities
    Dim Interval_Name() As String 'Name of processing intervals
    ReDim Mat_Name(1 To i)
    ReDim EUtility_Name(1 To ute)
    ReDim MUtility_Name(1 To utm)
    ReDim Interval_Name(1 To kk)


'[5] System Size ReDim based on Interval Number
    Dim n_feed As Long 'Number of raw material intervals
    Dim n_proc As Long 'Number of processing intervals
    Dim n_prod As Long 'Number of product intervals
    n_feed = Worksheets("S3").Range("F13").Value
    n_prod = Worksheets("S3").Cells(Rows.Count, "F").End(xlUp).Value
    n_proc = kk - n_feed - n_prod
    

'[6] Define Process Variables
    '[MASS][FEEDSTOCK SPECIFICATION]
    Dim phi() As Double              '[ARRAY]   Array of size i for each kk (with only kk up to n_feed filled) composition of raw material component i in feedstock interval kk
    
    '[MASS][INCOMING STREAMS INTO INTERVAL KK]
    Dim F() As Double                '[MULTI]   Flow of compound i (INCOMING) from interval k to interval kk. Array of size i for each kk
    Dim F_TOTAL() As Double          '[ARRAY]   Total Flow between source k and destination kk. Array of size kk for each k
    Dim F_in() As Double             '[ARRAY]   Flow of compound i entering interval kk. Array of size i for each kk F_in(kk,i) = ¢²F(k,kk,i)
    Dim F_in_TOTAL() As Double       '[SCALAR]  Total mass flow entering interval kk. F_in_TOTAL(kk) = ¢²F_in(kk,i)
   
    '[MASS][RAW MATERIAL SPECIFICATION]
    Dim R_B() As Double              '[ARRAY]   Flow of basis material entering Raw-Mat Mixing. It is a binary array of size i for each kk with 1 non-zero entry corresponding to the Basis Material
    Dim R_SA() As Double             '[ARRAY]   Array of size i for each kk corresponding to specific raw material addition values for compound i
    Dim F_RB() As Double             '[ARRAY]   BASIS MATERIAL mass flow for Raw Material Specific Addition. F_RB(kk,i) = F_in(kk,i) * R_B(kk,i)
    Dim F_RB_TOTAL() As Double       '[SCALAR]  The mass flow of the basis material for each interval k. F_RB_TOTAL(kk) = ¢²F_RB(kk,i)
    Dim R_M() As Double              '[ARRAY]   Flow of compound i entering interval kk as RAW MATERIAL. R_M(kk,i) = R_SA(kk,i) * F_RB_TOTAL(kk)
    Dim F_M() As Double              '[ARRAY]   Flow of compound i after raw material mix in interval kk. F_M(kk,i) = F_in(kk,i) + R_M(kk,i)
    Dim F_M_TOTAL() As Double        '[SCALAR]  Total mass flow after raw material mixing in interval kk. F_M_TOTAL(kk) = ¢²F_M(kk,i)
    
    '[MASS][REACTION]
    Dim FC_X() As Double             '[ARRAY]   Array of size i for each kk with 1 non-zero entry corresponding to the FRACTIONAL CONVERSION of the Key Reacting Material
    Dim K_X() As Double              '[ARRAY]   Array of size i for each kk with 1 non-zero entry being the reacted mass of the Key Reacting Material. K_X(kk,i) = FC_X(kk,i) .* F_M(kk,i)
    Dim K_X_TOTAL() As Double        '[SCALAR]  Total mass flow of the Key Reacting Component. K_X_TOTAL(kk) = ¢²K_X(kk,i)
    Dim SC_X() As Double             '[ARRAY]   Array of size i for each kk corresponding to the specific consumption of non-key reacting components
    Dim NK_X() As Double             '[ARRAY]   Array of size i for each kk corresponding to the reacted mass of the Non-Key Reacting Materials. NK_X(kk,i) = SC_X(kk,i) * K_X_TOTAL(kk)
    Dim NK_X_TOTAL() As Double       '[SCALAR]  Total mass flow of all Non-Key Reacting Components. NK_X_TOTAL(kk) = ¢²NK_X(kk,i)
    Dim M_X_TOTAL() As Double        '[SCALAR]  Total reacted mass at interval kk. M_X_TOTAL(kk) = K_X_TOTAL(kk) + NK_X_TOTAL(kk)
    Dim FY_X() As Double             '[ARRAY]   Array of size i for each kk corrsponding to the fractional mass-yields of product.
    Dim P_X() As Double              '[ARRAY]   Flow of compound i produced from the reaction in interval kk. P_X(kk,i) = FY_X(kk,i) * M_X_TOTAL(kk)
    Dim F_RX() As Double             '[ARRAY]   Flow of compound i after reaction from interval kk. F_RX(kk,i) = F_M(kk,i) - K_X(kk,i) - NK_X(kk,i) + P_X(kk,i)
    Dim F_RX_TOTAL() As Double       '[SCALAR]  Total mass flow after reaction in interval kk. F_RX_TOTAL(kk) = ¢²F_RX(kk,i)

    '[MASS][WASTE PURGE]
    Dim WP_Frac() As Double          '[ARRAY]   Array of size i for each kk corresponding to fraction of component i purged as waste
    Dim W() As Double                '[ARRAY]   Purged waste flow of compound i out of interval kk. W(kk,i) = F_RX(kk,i) - F_W(kk,i)
    Dim F_W() As Double              '[ARRAY]   Flow of compound i after waste separation in interval kk. F_W(kk,i) = F_RX(kk,i) * (1 - WP_Frac(kk,i))
    Dim F_W_TOTAL() As Double        '[SCALAR]  Total mass flow after waste separation in interval kk. F_W_TOTAL(kk) = ¢²F_W(kk,i)
    
    '[MASS][SEPARATION]
    Dim Sep_Frac() As Double         '[ARRAY]   Array of size i for each kk corresponding to fraction of component i in primary separator. Default values for Sep_Frac are 1 for all i, meaning no separation by default
    Dim F_OUT1() As Double           '[ARRAY]   Flow of compound i in primary outlet of separation in interval kk. F_OUT1(kk,i) = Sep_Frac(kk,i) * F_W(kk,i)
    Dim F_OUT2() As Double           '[ARRAY]   Flow of compound i in secondary outlet of separation in interval kk. F_OUT2(kk,i) = F_W(kk,i) - F_OUT1(kk,i)
    Dim F_1() As Double              '[MULTI]   Flow of compound i from interval k to interval kk (primary). If the current interval has only 1 connection to the next interval, there is no stream split of the primary outlet and hence F_1() = F_OUT1()
    Dim F_2() As Double              '[MULTI]   Flow of compound i from interval k to interval kk (secondary)
   
    '[UTILITIES][ENERGY UTILITY CONSUMPTION]
    Dim mu_RM_ENERGY() As Double            '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Raw Material Mixing
    Dim mu_RX_ENERGY() As Double            '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Reaction
    Dim mu_WS_ENERGY() As Double            '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Waste Purge
    Dim mu_PS_ENERGY() As Double            '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Product Separation
    Dim EU_RM() As Double                   '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for RM Mixing. EU_RM(kk,ute) = mu_RM_ENERGY(kk,ute) * F_M_TOTAL(kk)
    Dim EU_RX() As Double                   '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Reaction. EU_RX(kk,ute) = mu_RX_ENERGY(kk,ute) * F_RX_TOTAL(kk)
    Dim EU_WS() As Double                   '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Waste Purge. EU_WS(kk,ute) = mu_WS_ENERGY(kk,ute) * F_W_TOTAL(kk)
    Dim EU_PS() As Double                   '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Product Separation. EU_PS(kk,ute) = mu_PS_ENERGY(kk,ute) * F_W_TOTAL(kk)
    Dim EU() As Double                      '[ARRAY] Array of size ute for each kk. Total ENERGY UTILITY FLOWS for interval kk. EU(kk,ute) = EU_RM(kk,ute) + EU_RX(kk,ute) + EU_WS(kk,ute) + EU_PS(kk,ute)

    '[UTILITIES][MASS UTILITY CONSUMPTION]
    Dim mu_RM_MASS() As Double              '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Raw Material Mixing
    Dim mu_RX_MASS() As Double              '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Reaction
    Dim mu_WS_MASS() As Double              '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Waste Purge
    Dim mu_PS_MASS() As Double              '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Product Separation
    Dim MU_RM() As Double                   '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for RM Mixing. MU_RM(kk,utm) = mu_RM_MASS(kk,utm) * F_M_TOTAL(kk)
    Dim MU_RX() As Double                   '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Reaction. MU_RX(kk,utm) = mu_RX_MASS(kk,utm) * F_RX_TOTAL(kk)
    Dim MU_WS() As Double                   '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Waste Purge. MU_WS(kk,utm) = mu_WS_MASS(kk,utm) * F_W_TOTAL(kk)
    Dim MU_PS() As Double                   '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Product Separation. MU_PS(kk,utm) = mu_PS_MASS(kk,utm) * F_W_TOTAL(kk)
    Dim MU() As Double                      '[ARRAY] Array of size utm for each kk. Total MASS UTILITY FLOWS for interval kk. MU(kk,utm) = MU_RM(kk,utm) + MU_RX(kk,utm) + MU_WS(kk,utm) + MU_PS(kk,utm)
    
    '[TRANSPORTATION][TRANSPORTATION MATRIX]
    Dim D_1() As Double                     '[MULTI] Transportation distance from interval k to kk via transportation method tp for primary stream. Size of kxkk for each tp
    Dim D_2() As Double                     '[MULTI] Transportation distance from interval k to kk via transportation method tp for secondary stream. Size of kxkk for each tp
    Dim D_1_TOTAL() As Double               '[SCALAR] Total distance in the entire process for each transportation mode tp. Size of 1xtp
    Dim D_2_TOTAL() As Double               '[SCALAR] Total distance in the entire process for each transportation mode tp. Size of 1xtp
    

'[7] ReDim Process Variables based on System Size
    '[MASS][FEEDSTOCK SPECIFICATION]
    ReDim phi(1 To kk, 1 To i) As Double                          '[ARRAY]   Array of size i for each kk (with only kk up to n_feed filled) composition of raw material component i in feedstock interval kk
    
    '[MASS][INCOMING STREAMS INTO INTERVAL KK]
    ReDim F(1 To k, 1 To kk, 1 To i) As Double                    '[MULTI]   Flow of compound i (INCOMING) from interval k to interval kk. Array of size i for each kk
    ReDim F_TOTAL(1 To k, 1 To kk) As Double                      '[ARRAY]   Total Flow between source k and destination kk. Array of size kk for each k
    ReDim F_in(1 To kk, 1 To i) As Double                         '[ARRAY]   Flow of compound i entering interval kk. Array of size i for each kk F_in(kk,i) = ¢²F(k,kk,i)
    ReDim F_in_TOTAL(1 To kk) As Double                           '[SCALAR]  Total mass flow entering interval kk. F_in_TOTAL(kk) = ¢²F_in(kk,i)
   
    '[MASS][RAW MATERIAL SPECIFICATION]
    ReDim R_B(1 To kk, 1 To i) As Double                          '[ARRAY]   Flow of basis material entering Raw-Mat Mixing. It is a binary array of size i for each kk with 1 non-zero entry corresponding to the Basis Material
    ReDim R_SA(1 To kk, 1 To i) As Double                         '[ARRAY]   Array of size i for each kk corresponding to specific raw material addition values for compound i
    ReDim F_RB(1 To kk, 1 To i) As Double                         '[ARRAY]  BASIS MATERIAL mass flow for Raw Material Specific Addition. F_RB(kk,i) = F_in(kk,i) * R_B(kk,i)
    ReDim F_RB_TOTAL(1 To kk) As Double                           '[SCALAR]  The mass flow of the basis material for each interval k. F_RB_TOTAL(kk) = ¢²F_RB(kk,i)
    ReDim R_M(1 To kk, 1 To i) As Double                          '[ARRAY]   Flow of compound i entering interval kk as RAW MATERIAL. R_M(kk,i) = R_SA(kk,i) * F_RB_TOTAL(kk)
    ReDim F_M(1 To kk, 1 To i) As Double                          '[ARRAY]   Flow of compound i after raw material mix in interval kk. F_M(kk,i) = F_in(kk,i) + R_M(kk,i)
    ReDim F_M_TOTAL(1 To kk) As Double                            '[SCALAR]  Total mass flow after raw material mixing in interval kk. F_M_TOTAL(kk) = ¢²F_M(kk,i)
    
    '[MASS][REACTION]
    ReDim FC_X(1 To kk, 1 To i) As Double                         '[ARRAY]   Array of size i for each kk with 1 non-zero entry corresponding to the FRACTIONAL CONVERSION of the Key Reacting Material
    ReDim K_X(1 To kk, 1 To i) As Double                          '[ARRAY]   Array of size i for each kk with 1 non-zero entry being the reacted mass of the Key Reacting Material. K_X(kk,i) = FC_X(kk,i) .* F_M(kk,i)
    ReDim K_X_TOTAL(1 To kk) As Double                            '[SCALAR]  Total mass flow of the Key Reacting Component. K_X_TOTAL(kk) = ¢²K_X(kk,i)
    ReDim SC_X(1 To kk, 1 To i) As Double                         '[ARRAY]   Array of size i for each kk corresponding to the specific consumption of non-key reacting components
    ReDim NK_X(1 To kk, 1 To i) As Double                         '[ARRAY]   Array of size i for each kk corresponding to the reacted mass of the Non-Key Reacting Materials. NK_X(kk,i) = SC_X(kk,i) * K_X_TOTAL(kk)
    ReDim NK_X_TOTAL(1 To kk) As Double                           '[SCALAR]  Total mass flow of all Non-Key Reacting Components. NK_X_TOTAL(kk) = ¢²NK_X(kk,i)
    ReDim M_X_TOTAL(1 To kk) As Double                            '[SCALAR]  Total reacted mass at interval kk. M_X_TOTAL(kk) = K_X_TOTAL(kk) + NK_X_TOTAL(kk)
    ReDim FY_X(1 To kk, 1 To i) As Double                         '[ARRAY]   Array of size i for each kk corrsponding to the fractional mass-yields of product.
    ReDim P_X(1 To kk, 1 To i) As Double                          '[ARRAY]   Flow of compound i produced from the reaction in interval kk. P_X(kk,i) = FY_X(kk,i) * M_X_TOTAL(kk)
    ReDim F_RX(1 To kk, 1 To i) As Double                         '[ARRAY]   Flow of compound i after reaction from interval kk. F_RX(kk,i) = F_M(kk,i) - K_X(kk,i) - NK_X(kk,i) + P_X(kk,i)
    ReDim F_RX_TOTAL(1 To kk) As Double                           '[SCALAR]  Total mass flow after reaction in interval kk. F_RX_TOTAL(kk) = ¢²F_RX(kk,i)

    '[MASS][WASTE PURGE]
    ReDim WP_Frac(1 To kk, 1 To i) As Double                      '[ARRAY]   Array of size i for each kk corresponding to fraction of component i purged as waste
    ReDim W(1 To kk, 1 To i) As Double                            '[ARRAY]   Purged waste flow of compound i out of interval kk. W(kk,i) = F_RX(kk,i) - F_W(kk,i)
    ReDim F_W(1 To kk, 1 To i) As Double                          '[ARRAY]   Flow of compound i after waste separation in interval kk. F_W(kk,i) = F_RX(kk,i) * (1 - WP_Frac(kk,i))
    ReDim F_W_TOTAL(1 To kk) As Double                            '[SCALAR]  Total mass flow after waste separation in interval kk. F_W_TOTAL(kk) = ¢²F_W(kk,i)
    
    '[MASS][SEPARATION]
    ReDim Sep_Frac(1 To kk, 1 To i) As Double                     '[ARRAY]   Array of size i for each kk corresponding to fraction of component i in primary separator. Default values for Sep_Frac are 1 for all i, meaning no separation by default
    ReDim F_OUT1(1 To kk, 1 To i) As Double                       '[ARRAY]   Flow of compound i in primary outlet of separation in interval kk. F_OUT1(kk,i) = Sep_Frac(kk,i) * F_W(kk,i)
    ReDim F_OUT2(1 To kk, 1 To i) As Double                       '[ARRAY]   Flow of compound i in secondary outlet of separation in interval kk. F_OUT2(kk,i) = F_W(kk,i) - F_OUT1(kk,i)
    ReDim F_1(1 To k, 1 To kk, 1 To i) As Double                  '[MULTI]   Flow of compound i from interval k to interval kk (primary). If the current interval has only 1 connection to the next interval, there is no stream split of the primary outlet and hence F_1() = F_OUT1()
    ReDim F_2(1 To k, 1 To kk, 1 To i) As Double                  '[MULTI]   Flow of compound i from interval k to interval kk (secondary)
   
    '[UTILITIES][ENERGY UTILITY CONSUMPTION]
    ReDim mu_RM_ENERGY(1 To kk, 1 To ute) As Double               '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Raw Material Mixing
    ReDim mu_RX_ENERGY(1 To kk, 1 To ute) As Double               '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Reaction
    ReDim mu_WS_ENERGY(1 To kk, 1 To ute) As Double               '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Waste Purge
    ReDim mu_PS_ENERGY(1 To kk, 1 To ute) As Double               '[ARRAY] Array of size ute for each kk corresponding to specific consumption of ENERGY UTILITIES for Product Separation
    ReDim EU_RM(1 To kk, 1 To ute) As Double                      '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for RM Mixing. EU_RM(kk,ute) = mu_RM_ENERGY(kk,ute) * F_M_TOTAL(kk)
    ReDim EU_RX(1 To kk, 1 To ute) As Double                      '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Reaction. EU_RX(kk,ute) = mu_RX_ENERGY(kk,ute) * F_RX_TOTAL(kk)
    ReDim EU_WS(1 To kk, 1 To ute) As Double                      '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Waste Purge. EU_WS(kk,ute) = mu_WS_ENERGY(kk,ute) * F_RX_TOTAL(kk)
    ReDim EU_PS(1 To kk, 1 To ute) As Double                      '[ARRAY] Array of size ute for each kk >> ENERGY UTILITY FLOWS for Product Separation. EU_PS(kk,ute) = mu_PS_ENERGY(kk,ute) * F_W_TOTAL(kk)
    ReDim EU(1 To kk, 1 To ute) As Double                         '[ARRAY] Array of size ute for each kk. Total ENERGY UTILITY FLOWS for interval kk. EU(kk,ute) = EU_RM(kk,ute) + EU_RX(kk,ute) + EU_WS(kk,ute) + EU_PS(kk,ute)

    '[UTILITIES][MASS UTILITY CONSUMPTION]
    ReDim mu_RM_MASS(1 To kk, 1 To utm) As Double                 '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Raw Material Mixing
    ReDim mu_RX_MASS(1 To kk, 1 To utm) As Double                 '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Reaction
    ReDim mu_WS_MASS(1 To kk, 1 To utm) As Double                 '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Waste Purge
    ReDim mu_PS_MASS(1 To kk, 1 To utm) As Double                 '[ARRAY] Array of size utm for each kk corresponding to specific consumption of MASS UTILITIES for Product Separation
    ReDim MU_RM(1 To kk, 1 To utm) As Double                      '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for RM Mixing. MU_RM(kk,utm) = mu_RM_MASS(kk,utm) * F_M_TOTAL(kk)
    ReDim MU_RX(1 To kk, 1 To utm) As Double                      '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Reaction. MU_RX(kk,utm) = mu_RX_MASS(kk,utm) * F_RX_TOTAL(kk)
    ReDim MU_WS(1 To kk, 1 To utm) As Double                      '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Waste Purge. MU_WS(kk,utm) = mu_WS_MASS(kk,utm) * F_RX_TOTAL(kk)
    ReDim MU_PS(1 To kk, 1 To utm) As Double                      '[ARRAY] Array of size utm for each kk >> MASS UTILITY FLOWS for Product Separation. MU_PS(kk,utm) = mu_PS_MASS(kk,utm) * F_W_TOTAL(kk)
    ReDim MU(1 To kk, 1 To utm) As Double                         '[ARRAY] Array of size utm for each kk. Total MASS UTILITY FLOWS for interval kk. MU(kk,utm) = MU_RM(kk,utm) + MU_RX(kk,utm) + MU_WS(kk,utm) + MU_PS(kk,utm)

    '[TRANSPORTATION][TRANSPORTATION MATRIX]
    ReDim D_1(1 To k, 1 To kk, 1 To tp) As Double                 '[MULTI] Transportation distance from interval k to kk via transportation method tp for primary stream. Size of kxkk for each tp
    ReDim D_2(1 To k, 1 To kk, 1 To tp) As Double                 '[MULTI] Transportation distance from interval k to kk via transportation method tp for secondary stream. Size of kxkk for each tp
    ReDim D_1_TOTAL(1 To tp) As Double                            '[SCALAR] Total distance in the entire process for each transportation mode tp. Size of 1xtp
    ReDim D_2_TOTAL(1 To tp) As Double                            '[SCALAR] Total distance in the entire process for each transportation mode tp. Size of 1xtp


'[8] Binary variables
    Dim y_P() As Byte                            '[SUPERSTRUCTURE] Selection of process interval kk for primary stream
    Dim y_S() As Byte                            '[SUPERSTRUCTURE] Selection of process interval kk for secondary stream
    Dim xi_P() As Byte                           '[CURRENT PATHWAY] Connection of primary stream between interval k and interval kk
    Dim xi_S() As Byte                           '[CURRENT PATHWAY] Connection of secondary stream between interval k and interval kk
    ReDim y_P(1 To k, 1 To kk)                   '[ARRAY] Array of size k for each kk corresponding to whether or not a primary connection from k exists
    ReDim y_S(1 To k, 1 To kk)                   '[ARRAY] Array of size k for each kk corresponding to whether or not a secondary connection from k exists
    ReDim xi_P(1 To k, 1 To kk)                  '[ARRAY] Array of size k for each kk corresponding to a selected connection in current pathway
    ReDim xi_S(1 To k, 1 To kk)                  '[ARRAY] Array of size k for each kk corresponding to a selected connection in current pathway

'---------------------------------------------------------------------------------
'-----------------  [II] RECALL DATA AND ASSIGN TO VARIABLE ----------------------
'---------------------------------------------------------------------------------
'[9] Recall Inserted Data and Assign to Variables
    'RECALL COMPONENT NAMES
    For a = 1 To i
       Mat_Name(a) = Worksheets("B2").Range("C" & 3 + a).Value
    Next a
    For a = 1 To ute
      EUtility_Name(a) = Worksheets("B3").Range("C" & 4 + a).Value
    Next a
    For a = 1 To utm
      MUtility_Name(a) = Worksheets("B4").Range("C" & 4 + a).Value
    Next a
    For a = 1 To kk
      Interval_Name(a) = Worksheets("B10").Range("D" & 7 + a).Value
    Next a
   
    'RECALL PROCESS INPUT DATA: In other words, skip kk index to after n_feed and 1 to n_proc. Skip kk index of product intervals
    n = 1
    For a = n_feed + 1 To n_feed + n_proc
        For b = 1 To i
            'Raw Material
            R_B(a, b) = Worksheets("B10").Cells(7 + kk + 6 + n_feed + 10 + n, 3 + b).Value
            R_SA(a, b) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 10 + n_proc + n, 3 + b).Value
            'Reaction
            FC_X(a, b) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 20 + (2 * n_proc) + n, 3 + b).Value
            SC_X(a, b) = Worksheets("B10").Cells(7 + kk + 18 + n_feed + 20 + (3 * n_proc) + n, 3 + b).Value
            FY_X(a, b) = Worksheets("B10").Cells(7 + kk + 24 + n_feed + 20 + (4 * n_proc) + n, 3 + b).Value
            'Waste Purge
            WP_Frac(a, b) = Worksheets("B10").Cells(7 + kk + 30 + n_feed + 20 + (5 * n_proc) + n, 3 + b).Value
        Next b
        For bb = 1 To ute
            mu_RM_ENERGY(a, bb) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 10 + n_proc + n, 3 + i + bb).Value
            mu_RX_ENERGY(a, bb) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 20 + (2 * n_proc) + n, 3 + i + bb).Value
            mu_WS_ENERGY(a, bb) = Worksheets("B10").Cells(7 + kk + 30 + n_feed + 20 + (5 * n_proc) + n, 3 + i + bb).Value
        Next bb
        For bbb = 1 To utm
            mu_RM_MASS(a, bbb) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 10 + n_proc + n, 3 + i + ute + bbb).Value
            mu_RX_MASS(a, bbb) = Worksheets("B10").Cells(7 + kk + 12 + n_feed + 20 + (2 * n_proc) + n, 3 + i + ute + bbb).Value
            mu_WS_MASS(a, bbb) = Worksheets("B10").Cells(7 + kk + 30 + n_feed + 20 + (5 * n_proc) + n, 3 + i + ute + bbb).Value
        Next bbb
        n = n + 1
    Next a
    
    'RECALL FEEDSTOCK DATA: In other words, only consider the first feed intervals
    For a = 1 To n_feed
        F_in_TOTAL(a) = Worksheets("B10").Cells(7 + kk + 6 + a, 4 + i).Value
        For b = 1 To i
            phi(a, b) = Worksheets("B10").Cells(7 + kk + 6 + a, 3 + b).Value
        Next b
    Next a
   
   'RECALL PRODUCT SEPARATION DATA: In other words, consider all intervals including feed and product
    For a = 1 To kk
        For b = 1 To i
            Sep_Frac(a, b) = Worksheets("B10").Cells(7 + kk + 36 + n_feed + 20 + (6 * n_proc) + a, 3 + b).Value
        Next b
        For bb = 1 To ute
            mu_PS_ENERGY(a, bb) = Worksheets("B10").Cells(7 + kk + 36 + n_feed + 20 + (6 * n_proc) + a, 3 + i + bb).Value
        Next bb
        For bbb = 1 To utm
            mu_PS_MASS(a, bbb) = Worksheets("B10").Cells(7 + kk + 36 + n_feed + 20 + (6 * n_proc) + a, 3 + i + ute + bbb).Value
        Next bbb
    Next a
   
    'RECALL TRANSPORTATION AND INTERVAL CONNECTION INFO
    For a = 1 To k
        For b = 1 To kk
            For C = 1 To tp
                D_1(a, b, C) = Worksheets("B11").Cells((C * 6) + ((C - 1) * kk) + a, 3 + b).Value
                D_2(a, b, C) = Worksheets("B11").Cells((C * 6) + ((C - 1) * kk) + a, 3 + kk + b).Value
            Next C
            'All connections in superstructure
            y_P(a, b) = Worksheets("B7").Cells(7 + a, 3 + b).Value
            y_S(a, b) = Worksheets("B7").Cells(12 + kk + a, 3 + b).Value
            'Current connections for specified pathway
            xi_P(a, b) = Worksheets("B12").Cells(7 + a, 3 + b).Value
            xi_S(a, b) = Worksheets("B12").Cells(12 + kk + a, 3 + b).Value
        Next b
    Next a



'---------------------------------------------------------------------------------
'-----------------  [III] PROCESS INTERVAL EQUATIONS -----------------------------
'---------------------------------------------------------------------------------
'***** MASS BALANCE *****
'[10] Mass Balance Equations
For Count = 1 To steps
    '------- INCOMING FLOW -------
    '1a) [INCOMING FLOW] Flow of compound i entering interval kk from interval k
    'F_in(kk,i) = ¢²F(k,kk,i)
    x = 0
    For a = n_feed + 1 To kk
        For b = 1 To i
            For C = 1 To k
                x = x + F(C, a, b)              'The F() matrix is initialized as zero. For each current intervan and for each material, add to x for each source interval k (looping through C)
                F_in(a, b) = x                  'Continue looping through source interval k. Each time, the total component inlet into kk increases until the loop stops
            Next C
            x = 0
        Next b
      x = 0
    Next a
    
    '1b) [INCOMING FLOW] Total mass flow entering interval kk
    'F_in_TOTAL(kk) = ¢²F_in(kk,i)
    For a = n_feed + 1 To kk
        For b = 1 To i
            x = x + F_in(a, b)                 'Initialize x = 0. Add onto x the component flowrate (b) for each current interval (a)
            F_in_TOTAL(a) = x                  'Continue summing the component flowrate (b) for interval (a) until terminates (when b = i)
        Next b
        x = 0
    Next a


    '------- RAW MATERIAL -------
    '2a) [RAW MATERIAL] Raw Material Basis, Flow of compound i as basis material in interval kk
    'F_RB(kk,i) = F_in(kk,i) .* R_B(kk,i)
    For a = 1 To kk
        For b = 1 To i
            F_RB(a, b) = F_in(a, b) * R_B(a, b)
        Next b
    Next a

    '2b) [RAW MATERIAL] Raw Material Basis, Total Flow of Basis Material in Interval kk
    'F_RB_TOTAL(kk) = ¢²F_RB(kk,i)
    x = 0
    For a = 1 To kk
        For b = 1 To i
            x = x + F_RB(a, b)
            F_RB_TOTAL(a) = x
        Next b
        x = 0
    Next a
    
    '2c) [RAW MATERIAL] Raw Material Addition, Flowrate of Compound i in Interval kk as a Raw Material
    'R_M(kk,i) = R_SA(kk,i) * F_RB_TOTAL(kk)
    For a = 1 To kk
        For b = 1 To i
            R_M(a, b) = R_SA(a, b) * F_RB_TOTAL(a)
        Next b
    Next a

    '2d) [RAW MATERIAL] Flow after Raw Material Mixing. Flowrate of Compound i in Interval kk after RM Mixing
    'F_M(kk,i) = F_in(kk,i) + R_M(kk,i)
    For a = 1 To kk
        For b = 1 To i
            F_M(a, b) = F_in(a, b) + R_M(a, b)
        Next b
    Next a
    
    '2e) [RAW MATERIAL] Total Flow after Raw Material Mixing
    'F_M_TOTAL(kk) = ¢²F_M(kk,i)
    x = 0
    For a = 1 To kk
       For b = 1 To i
          x = x + F_M(a, b)
          F_M_TOTAL(a) = x
       Next b
       x = 0
    Next a
    
    
    '------- REACTION -------
    '3a) [REACTION] Reacted Mass of the Key Reacting Material (1 non-zero entry per interval kk)
    'K_X(kk,i) = FC_X(kk,i) * F_M(kk,i)
    For a = 1 To kk
        For b = 1 To i
            K_X(a, b) = FC_X(a, b) * F_M(a, b)
        Next b
    Next a
    
    '3b) [REACTION] Total Reacted Mass of the Key Reacting Material in Interval kk
    'K_X_TOTAL(kk) = ¢²K_X(kk,i)
    x = 0
    For a = 1 To kk
        For b = 1 To i
            x = x + K_X(a, b)
            K_X_TOTAL(a) = x
        Next b
        x = 0
    Next a

    '3c) [REACTION] Reacted Mass of Non-Key Reacting Materials, Component i for Interval kk
    'NK_X(kk,i) = SC_X(kk,i) * K_X_TOTAL(kk)
    For a = 1 To kk
        For b = 1 To i
            NK_X(a, b) = SC_X(a, b) * K_X_TOTAL(a)
        Next b
    Next a
    
    '3d) [REACTION] Total Reacted Mass of Non-Key Reacting Material in Interval kk
    'NK_X_TOTAL(kk) = ¢²NK_X(kk,i)
    x = 0
    For a = 1 To kk
        For b = 1 To i
            x = x + NK_X(a, b)
            NK_X_TOTAL(a) = x
        Next b
        x = 0
    Next a

    '3d) [REACTION] Total Reacted Mass in Interval kk
    'M_X_TOTAL(kk) = K_X_TOTAL(kk) + NK_X_TOTAL(kk)
    For a = 1 To kk
        M_X_TOTAL(a) = K_X_TOTAL(a) + NK_X_TOTAL(a)
    Next a
    
    '3e) [REACTION] Produced Mass Flowrate of Component i for Interval kk
    'P_X(kk,i) = FY_X(kk,i) * M_X_TOTAL(kk)
    For a = 1 To kk
        For b = 1 To i
            P_X(a, b) = FY_X(a, b) * M_X_TOTAL(a)
        Next b
    Next a

    '3f) [REACTION] After-Reaction Flowrate of Component i for Interval kk
    'F_RX(kk,i) = F_M(kk,i) - K_X(kk,i) - NK_X(kk,i) + P_X(kk,i)
    For a = 1 To kk
        For b = 1 To i
            F_RX(a, b) = F_M(a, b) - K_X(a, b) - NK_X(a, b) + P_X(a, b)
        Next b
    Next a
    
    '3g) [REACTION] Total After-Reaction Flowrate for Interval kk
    'F_RX_TOTAL(kk) = ¢²F_RX(kk,i)
    x = 0
    For a = 1 To kk
        For b = 1 To i
            x = x + F_RX(a, b)
            F_RX_TOTAL(a) = x
        Next b
        x = 0
    Next a
    
    
    '------- WASTE PURGE -------
    '4a) [WASTE PURGE] Post Waste Purge Flowrate of Component i for Interval kk
    'F_W(kk,i) = F_RX(kk,i) * (1 - WP_Frac(kk,i))
    For a = 1 To kk
        For b = 1 To i
            If a < n_feed + 1 Then
                F_W(a, b) = phi(a, b) * F_in_TOTAL(a)
            Else
                F_W(a, b) = F_RX(a, b) * (1 - WP_Frac(a, b))
            End If
        Next b
    Next a
    
    '4b) [WASTE PURGE] Flowrate of Compound i Removed as Waste for Interval kk
    'W(kk,i) = F_RX(kk,i) - F_W(kk,i)
    For a = 1 To kk
        For b = 1 To i
            If a > n_feed Then
                W(a, b) = F_RX(a, b) - F_W(a, b)
            End If
        Next b
    Next a

    '4c) [WASTE PURGE] Total Flowrate after Waste Separation for Interval kk
    'F_W_TOTAL(kk) = ¢²F_W(kk,i)
    x = 0
    For a = 1 To kk
        For b = 1 To i
            x = x + F_W(a, b)
            F_W_TOTAL(a) = x
        Next b
        x = 0
    Next a
    
    
    '------- PRODUCT SEPARATION -------
    '5a) [PRODUCT SEPARATION] Flowrate of Compound i in the Primary Stream after Separation in IntervaL KK
    'F_OUT1(kk,i) = Sep_Frac(kk,i) * F_W(kk,i)
    For a = 1 To kk
        For b = 1 To i
            F_OUT1(a, b) = F_W(a, b) * Sep_Frac(a, b)
        Next b
    Next a
    
    '5b) [PRODUCT SEPARATION] Flowrate of Compound i in the Secondary Stream after Separation in IntervaL KK
    'F_OUT2(kk,i) = F_W(kk,i) - F_OUT1(kk,i)
    For a = 1 To kk
        For b = 1 To i
            F_OUT2(a, b) = F_W(a, b) * (1 - Sep_Frac(a, b))
        Next b
    Next a
    
    
    '------- OUTGOING STREAM CONNECTION -------
    '6a) [CONNECTION] Connection of Primary and Secondary Streams between Interval k and kk
    x = 0
    For a = 1 To k
        For b = 1 To kk
            For C = 1 To i
                F_1(a, b, C) = F_OUT1(a, C) * xi_P(a, b) * y_P(a, b)
                F_2(a, b, C) = F_OUT2(a, C) * xi_S(a, b) * y_S(a, b)
            Next C
        Next b
    Next a

    '6b) [CONNECTION] Flow of Compound i entering Interval kk from Interval k via Primary/Secondary Streams
    For a = 1 To k
        For b = 1 To kk
            For C = 1 To i
                F(a, b, C) = F_1(a, b, C) + F_2(a, b, C)
            Next C
        Next b
    Next a
Next Count


'***** UTILITY & TRANSPORTATION *****
'[11] Energy Balance and Transportation Equations
'---- ENERGY UTILITIES ----
    For a = 1 To kk
        For b = 1 To ute
            EU_RM(a, b) = mu_RM_ENERGY(a, b) * F_M_TOTAL(a)
            EU_RX(a, b) = mu_RX_ENERGY(a, b) * F_RX_TOTAL(a)
            EU_WS(a, b) = mu_WS_ENERGY(a, b) * F_RX_TOTAL(a)
            EU_PS(a, b) = mu_PS_ENERGY(a, b) * F_W_TOTAL(a)
            EU(a, b) = EU_RM(a, b) + EU_RX(a, b) + EU_WS(a, b) + EU_PS(a, b)
        Next b
    Next a
    
    '---- MASS UTILITIES ----
    For a = 1 To kk
        For b = 1 To utm
            MU_RM(a, b) = mu_RM_MASS(a, b) * F_M_TOTAL(a)
            MU_RX(a, b) = mu_RX_MASS(a, b) * F_RX_TOTAL(a)
            MU_WS(a, b) = mu_WS_MASS(a, b) * F_RX_TOTAL(a)
            MU_PS(a, b) = mu_PS_MASS(a, b) * F_W_TOTAL(a)
            MU(a, b) = MU_RM(a, b) + MU_RX(a, b) + MU_WS(a, b) + MU_PS(a, b)
        Next b
    Next a
    
    '---- TRANSPORTATION ----
    'Calculate the Total Incoming Flow [F_in(a,b)] times the connection matrix for Distance
    x = 0
    For a = 1 To k
        For b = 1 To kk
            For C = 1 To i
                x = x + F(a, b, C)
            Next C
            F_TOTAL(a, b) = x
            x = 0
        Next b
        x = 0
    Next a
    
    'Multiply the Total Incoming Flow [F_in(a,b)] by the connection matrix for Distance
    y1 = 0
    y2 = 0
    For a = 1 To tp
        For b = 1 To k
            For C = 1 To kk
                y1 = y1 + (F_TOTAL(b, C) * D_1(b, C, a))
                y2 = y2 + (F_TOTAL(b, C) * D_2(b, C, a))
            Next C
        Next b
        D_1_TOTAL(a) = y1
        D_2_TOTAL(a) = y2
        y1 = 0
        y2 = 0
    Next a



'---------------------------------------------------------------------------------
'-----------------  [IV] GENERATE MASS BALANCE SPREADSHEET -----------------------
'---------------------------------------------------------------------------------
'[12] Generate Material Balance Tables
    ' "Materials" Dim Vars
    Dim cell_selected2 As Range
    Dim aa As Integer
    Dim Feed_Selected As Integer
    Dim num_flow_vars As Integer
    Dim is_first_interval As Integer
    Dim ii As Integer
    Dim jj As Integer
    Dim product_step As Integer
    Dim proc_count As Integer
    Dim ColInd As Integer
    
    ' "Materials" Header
    Set cell_selected2 = Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(2)
    With Range(cell_selected2, cell_selected2.Offset(1, 0))
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Material"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    ' "Materials" Listing Materials
    For aa = 1 To i
        With Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = Worksheets("B2").Cells(3 + aa, 3).Value
          .Interior.Color = RGB(221, 235, 247)
        End With
    Next aa
      
    ' "Intervals" Header: FEED INTERVALS
    Feed_Selected = Worksheets("S8").Range("F12").Value
    Set cell_selected2 = Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i - 1, 1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = Worksheets("B10").Cells(7 + Feed_Selected, 4).Value
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 1)
     .Value = "Feedrate, tons"
     .HorizontalAlignment = xlCenter
     .Interior.Color = RGB(221, 235, 247)
    End With

    ' "Intervals" Header: PROCESS INTERVALS- Define Sub-Headers
    is_first_interval = 1            ' In order to make the correct table offset
    num_flow_vars = 11               ' F_in, R_M, F_M, K_X, NK_X, P_X, F_RX, F_W, W, F_OUT1, F_OUT2
    Dim flow_vars(1 To 11)
    flow_vars(1) = "F_in"
    flow_vars(2) = "R_M"
    flow_vars(3) = "F_M"
    flow_vars(4) = "K_X"
    flow_vars(5) = "NK_X"
    flow_vars(6) = "P_X"
    flow_vars(7) = "F_RX"
    flow_vars(8) = "F_W"
    flow_vars(9) = "W"
    flow_vars(10) = "F_OUT1"
    flow_vars(11) = "F_OUT2"
    
    ' "Intervals" Header: PROCESS INTERVALS- Index Array for the Process + Product Intervals
    Dim Index_Array()
    ReDim Index_Array(1 To IntCount)
    For ii = 0 To IntCount - 1
        For jj = 1 To k
            If MainArray(ii) = Worksheets("B10").Cells(7 + jj, 4).Value Then
                Index_Array(ii + 1) = Worksheets("B10").Cells(7 + jj, 2).Value
            End If
        Next jj
    Next ii

    ' "Intervals" Header: PROCESS INTERVALS- Generating Tables
    Dim counter As Integer
    counter = 0
    For ii = 0 To IntCount - 1
        If Index_Array(ii + 1) <> steps Then
            Set cell_selected2 = Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i - 1, 2 + (counter * num_flow_vars))
            With Range(cell_selected2, cell_selected2.Offset(, num_flow_vars - 1))
               .Merge
               .VerticalAlignment = xlCenter
               .HorizontalAlignment = xlCenter
               .Value = MainArray(ii)
               .Font.Bold = True
               .Interior.Color = RGB(221, 235, 247)
            End With
            For a = 1 To num_flow_vars
               Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 2 + ((counter + 1) * num_flow_vars) - a).Value = flow_vars(12 - a)
               Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 2 + ((counter + 1) * num_flow_vars) - a).Interior.Color = RGB(221, 235, 247)
            Next a
            counter = counter + 1
        End If
    Next ii
    
    ' "Intervals" Header: PRODUCT- Generating Tables
    ColInd = Worksheets("O1").Cells(5, Columns.Count).End(xlToLeft).Column + num_flow_vars
    For ii = 0 To IntCount - 1
        If Index_Array(ii + 1) = steps Then
            Set cell_selected2 = Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i - 1, ColInd - 2)
            With cell_selected2
               .VerticalAlignment = xlCenter
               .HorizontalAlignment = xlCenter
               .Value = MainArray(ii)
               .Font.Bold = True
               .Interior.Color = RGB(221, 235, 247)
            End With
            Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i, ColInd - 2).Value = "Product Mass"
            Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Offset(-i, ColInd - 2).Interior.Color = RGB(221, 235, 247)
            ColInd = ColInd + 1
        End If
    Next ii

    ' Drawing cell boundaries
    Dim RowInd As Integer
    ColInd = Worksheets("O1").Cells(5, Columns.Count).End(xlToLeft).Column
    RowInd = Worksheets("O1").Cells(Rows.Count, "B").End(xlUp).Row
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O1").Range(Worksheets("O1").Cells(5, 2), Worksheets("O1").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    
    
'[13] Save Values into Table
    'Placeholder



'---------------------------------------------------------------------------------
'-----------------  [V] CALCULATION SUMMARY TABLES -------------------------------
'---------------------------------------------------------------------------------
'[14] Declare Variables
    Dim head As Range
    Dim temp As Integer
    Dim y As Integer
    Dim z As Integer
    
'[15] MATERIALS SUMMARY TABLE
    '[MATERIAL SUMMARY] Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With head
       .Value = "[1] MATERIAL BALANCE SUMMARY"
       .Font.Bold = True
       .HorizontalAlignment = xlLeft
    End With

    '[MATERIAL SUMMARY] Generate Table
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Material"
       .Interior.Color = RGB(221, 235, 247)
    End With
    For aa = 1 To i
        With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = Worksheets("B2").Cells(3 + aa, 3).Value
          .Interior.Color = RGB(221, 235, 247)
        End With
    Next aa
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Mass as Feedstock (tons/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 2)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Mass Added as RM (tons/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 3)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Mass Purged as Waste (tons/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-i, 4)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Mass as Product (tons/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    '[MATERIAL SUMMARY] Draw Lines
    ColInd = Worksheets("O2").Cells(5, Columns.Count).End(xlToLeft).Column
    RowInd = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Row
    temp = RowInd
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(5, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter


'[16] UTILITIES SUMMARY TABLE
    '[UTILITIES] Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(3)
    With head
       .Value = "[2] UTILITIES SUMMARY"
       .Font.Bold = True
       .HorizontalAlignment = xlLeft
    End With

    '[ENERGY UTILITIES] Sub-Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With head
       .Value = "Energy Utilities"
       .HorizontalAlignment = xlLeft
    End With
    
    '[ENERGY UTILITIES] Generate Table
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "E. Utilities"
       .Interior.Color = RGB(221, 235, 247)
    End With
    For aa = 1 To ute
        With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = Worksheets("B3").Cells(4 + aa, 3).Value
          .Interior.Color = RGB(221, 235, 247)
        End With
    Next aa
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-ute, 1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Total Consumed (GJ/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    '[ENERGY UTILITIES] Draw Lines
    ColInd = Worksheets("O2").Cells(temp + 5, Columns.Count).End(xlToLeft).Column
    RowInd = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Row
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 5, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    temp = RowInd

    '[MASS UTILITIES] Sub-Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With head
       .Value = "Mass Utilities"
       .HorizontalAlignment = xlLeft
    End With
    
    '[MASS UTILITIES] Generate Table
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "M. Utilities"
       .Interior.Color = RGB(221, 235, 247)
    End With
    For aa = 1 To utm
        With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = Worksheets("B4").Cells(4 + aa, 3).Value
          .Interior.Color = RGB(221, 235, 247)
        End With
    Next aa
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-utm, 1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Total Consumed (tons/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    '[ENERGY UTILITIES] Draw Lines
    ColInd = Worksheets("O2").Cells(temp + 2, Columns.Count).End(xlToLeft).Column
    RowInd = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Row
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 2, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    temp = RowInd


'[17] TRANSPORTATIONS SUMMARY TABLE
    '[TRANSPORTATIONS] Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(3)
    With head
       .Value = "[3] TRANSPORTATION SUMMARY"
       .Font.Bold = True
       .HorizontalAlignment = xlLeft
    End With
    
    '[TRANSPORTATIONS] Generate Table
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Transportations"
       .Interior.Color = RGB(221, 235, 247)
    End With
    For aa = 1 To tp
        With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = Worksheets("B5").Cells(4 + aa, 3).Value
          .Interior.Color = RGB(221, 235, 247)
        End With
    Next aa
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-tp, 1)
    With cell_selected2
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Total Mass-Displacement (km-ton/batch)"
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    '[TRANSPORTATIONS] Draw Lines
    ColInd = Worksheets("O2").Cells(temp + 4, Columns.Count).End(xlToLeft).Column
    RowInd = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Row
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    temp = RowInd


'[18] PRODUCT INTERVAL SUMMARY TABLE
    '[PRODUCT] Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(3)
    With head
       .Value = "[4] PRODUCT INTERVAL SUMMARY"
       .Font.Bold = True
       .HorizontalAlignment = xlLeft
    End With

    '[PRODUCT] Product Intervals Table Index
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With Range(cell_selected2, cell_selected2.Offset(, 1))
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Index"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    With cell_selected2.Offset(1)
       .Value = "Product Step"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
    With cell_selected2
       .Value = "Interval"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    ' [PRODUCT] Add a column for the total product rate
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
    With Range(cell_selected2, cell_selected2.Offset(, i - 1))
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Product Masses (ton)"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    For a = 1 To i
       With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
        .Value = Worksheets("B2").Range("C" & 3 + a).Value
        .Interior.Color = RGB(221, 235, 247)
       End With
    Next a

    ' [PRODUCT] Total Product Rate
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + i)
    With cell_selected2
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Product Interval"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, i + 1)
     .Value = "Total Product Mass (tons)"
     .Interior.Color = RGB(221, 235, 247)
    End With

    ' [PRODUCT] Product Intervals Numbering Index
    Dim prod_step As Integer
    Dim prod_interval As Integer
    For a = steps To steps
       prod_step = Worksheets("S3").Range("E" & 12 + a).Value
       prod_interval = Worksheets("S3").Range("F" & 12 + a).Value
       For b = 1 To prod_interval
          With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
            .Value = prod_step
            .Interior.Color = RGB(221, 235, 247)
          End With
          With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(1)
            .Value = b
            .Interior.Color = RGB(221, 235, 247)
          End With
       Next b
    Next a
    
    '[PRODUCT] Draw Lines
    ColInd = Worksheets("O2").Cells(temp + 5, Columns.Count).End(xlToLeft).Column
    RowInd = temp + 5 + prod_interval
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    temp = RowInd

'[18] TOTAL MASS SUMMARY (FOR EQUIPMENT COST SCALING)
    '[TOTAL MASS] Title
    Set head = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(3)
    With head
       .Value = "[5] INTERVAL MASS FLOW SUMMARY"
       .Font.Bold = True
       .HorizontalAlignment = xlLeft
    End With

    '[TOTAL MASS] Intervals Table Index
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
    With Range(cell_selected2, cell_selected2.Offset(, 1))
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Index"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    With cell_selected2.Offset(1)
       .Value = "Product Step"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
    With cell_selected2
       .Value = "Interval"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    
    ' [TOTAL MASS] Add a Columns for Total Flow Variables
    Set cell_selected2 = Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
    With Range(cell_selected2, cell_selected2.Offset(, 4 - 1))
       .Merge
       .VerticalAlignment = xlCenter
       .HorizontalAlignment = xlCenter
       .Value = "Total Mass Flows per Interval (tons)"
       .Font.Bold = True
       .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, 1)
     .Value = "Total Incoming (F_in_TOTAL)"
     .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, 2)
     .Value = "Total After RM-Mixing (F_M_TOTAL)"
     .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, 3)
     .Value = "Total After RXN (F_RX_TOTAL)"
     .Interior.Color = RGB(221, 235, 247)
    End With
    With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(, 4)
     .Value = "Total After Waste Purge (F_W_TOTAL)"
     .Interior.Color = RGB(221, 235, 247)
    End With

    ' [PRODUCT] Product Intervals Numbering Index
    Dim process_step As Integer
    Dim process_interval As Integer
    For a = 1 To steps - 2
       process_step = Worksheets("S3").Range("E" & 13 + a).Value
       process_interval = Worksheets("S3").Range("F" & 13 + a).Value
       For b = 1 To process_interval
          With Worksheets("O2").Cells(Rows.Count, "B").End(xlUp).Offset(1)
            .Value = process_step
            .Interior.Color = RGB(221, 235, 247)
          End With
          With Worksheets("O2").Cells(Rows.Count, "C").End(xlUp).Offset(1)
            .Value = b
            .Interior.Color = RGB(221, 235, 247)
          End With
       Next b
    Next a
    
    '[PRODUCT] Draw Lines
    ColInd = Worksheets("O2").Cells(temp + 5, Columns.Count).End(xlToLeft).Column
    RowInd = temp + 5 + n_proc
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).HorizontalAlignment = xlCenter
    Worksheets("O2").Range(Worksheets("O2").Cells(temp + 4, 2), Worksheets("O2").Cells(RowInd, ColInd)).VerticalAlignment = xlCenter
    
    
'[20] SAVE VALUES TO TABLES
    '***Saving Material Balance Values***
    'Mass as Feedstock
    For y = 1 To i
        Worksheets("O2").Cells(5 + y, 3).Value = F_W(Feed_Selected, y)
    Next y
    'Mass added as RM
    x = 0
    For y = 1 To i
        For z = 1 To kk
            x = x + R_M(z, y)
        Next z
        Worksheets("O2").Cells(5 + y, 4).Value = x
        x = 0
    Next y
    'Mass Purged as Waste
    x = 0
    For y = 1 To i
        For z = 1 To kk
            x = x + W(z, y)
        Next z
        Worksheets("O2").Cells(5 + y, 5).Value = x
        x = 0
    Next y
    'Mass as Product
    x = 0
    Dim starting_prod As Integer
    starting_prod = kk - n_prod
    For y = 1 To i
        For z = starting_prod + 1 To kk
            x = x + F_in(z, y)
        Next z
        Worksheets("O2").Cells(5 + y, 6).Value = x
        x = 0
    Next y

    '***Saving Utility Values***
    'Energy Utilities
    x = 0
    For y = 1 To ute
        For z = 1 To kk
            x = x + EU(z, y)
        Next z
        Worksheets("O2").Cells(10 + i + y, 3).Value = x
        x = 0
    Next y
    'Mass Utilities
    x = 0
    For y = 1 To utm
        For z = 1 To kk
            x = x + MU(z, y)
        Next z
        Worksheets("O2").Cells(12 + i + ute + y, 3).Value = x
        x = 0
    Next y

    '***Saving Transportation Values***
    For y = 1 To tp
        Worksheets("O2").Cells(16 + i + ute + utm + y, 3).Value = D_1_TOTAL(y) + D_2_TOTAL(y)
    Next y
    
    '***Save Product Masses by Interval***
    x = 0
    For z = 1 To prod_interval
        For y = 1 To i
            Worksheets("O2").Cells(21 + i + ute + utm + tp + z, 3 + y).Value = F_in(z + starting_prod, y)
            x = x + F_in(z + starting_prod, y)
        Next y
        Worksheets("O2").Cells(21 + i + ute + utm + tp + z, 4 + i).Value = x
        x = 0
    Next z

    '***Save Interval Total Masses***
    For z = 1 To n_proc
        Worksheets("O2").Cells(26 + i + ute + utm + tp + n_prod + z, 4).Value = F_in_TOTAL(n_feed + z)
        Worksheets("O2").Cells(26 + i + ute + utm + tp + n_prod + z, 5).Value = F_M_TOTAL(n_feed + z)
        Worksheets("O2").Cells(26 + i + ute + utm + tp + n_prod + z, 6).Value = F_RX_TOTAL(n_feed + z)
        Worksheets("O2").Cells(26 + i + ute + utm + tp + n_prod + z, 7).Value = F_W_TOTAL(n_feed + z)
    Next z
    

'---------------------------------------------------------------------------------
'-----------------  [VI] COMPUTATIONAL TIME DISPLAY ------------------------------
'---------------------------------------------------------------------------------
    'Mass Balance Checksum = 1
    Worksheets("O2").Cells(2, 6).Value = 1
    Worksheets("O1").Cells(2, 6).Value = 1
    MsgBox "Computation time : " & Format(Timer - Start, "0.000") & " seconds", vbExclamation, "TIPEM- Notice"
    Exit Sub

    'Error Handler
message:
    MsgBox "There was an error. Mass Balances not Calculated", vbExclamation, "TIPEM- Error"
End Sub






















' [MASS BALANCE MODEL] Create Interval Specification Table
Public Function TIPEM_Create_IntervalSpecTable()
' [1] Define Variables
    Dim n_step As Integer
    Dim n_comp As Integer
    Dim n_Eutil As Integer
    Dim n_Mutil As Integer
    Dim n_rxn As Integer
    Dim n_total_interval As Integer
    Dim n_feed_interval As Integer
    Dim n_proc_interval As Integer
    Dim n_prod_interval As Integer
    
    Dim n_step_interval As Integer
    Dim n_interval As Integer
    
    Dim a As Integer
    Dim b As Integer
    Dim n As Integer
    Dim x As Integer
    
    Dim head As Range
    Dim cell_selected As Range
    Dim cell_selected_2 As Range
    Dim cell_selected_3 As Range


'====================================================================
' [2] Specify and Assign System Variables
    Worksheets("B10").Activate
    Worksheets("B10").Cells(4, "B").Select
    n_step = Worksheets("S3").Range("H12").Value + 2
    n_comp = Worksheets("B2").Range("K3").Value
    n_Eutil = Worksheets("B3").Range("C1").Value
    n_Mutil = Worksheets("B4").Range("C1").Value
    n_total_interval = Worksheets("S3").Range("H14").Value
    n_feed_interval = Worksheets("S3").Range("F13").Value
    n_prod_interval = Worksheets("S3").Range("F" & 12 + n_step).Value
    n_proc_interval = n_total_interval - n_feed_interval - n_prod_interval


'====================================================================
' [3a] Interval Name Specification
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1)
      With head
         .Value = "[1] INTERVAL NAME SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      ' Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' Specify Name of Each Feedstock/Process/Product Interval
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(1))
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Name"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' Process Intervals Numbering Index
      For a = 1 To n_step
         n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
         n_interval = Worksheets("S3").Range("F" & 12 + a).Value
         For b = 1 To n_interval
            With Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Worksheets("B10").Cells(Rows.Count, "D").End(xlUp).Offset(1)
              .Value = "Placeholder"
            End With
         Next b
      Next a
      
      ' Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter


'====================================================================
' [3b] FEEDSTOCK Step Specification
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "[2] FEEDSTOCK SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      ' Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Feedstock Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' Feedstock Composition
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Feedstock Composition"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' Feedstock Feed Rate Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp)
      With Range(cell_selected, cell_selected.Offset(, 0))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Feedstock"
         .Font.Bold = True
         .Interior.Color = RGB(32, 196, 132)
      End With
      For a = 1 To 1
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp)
          .Value = "Feedrate"
          .Interior.Color = RGB(32, 196, 132)
         End With
      Next a
       
      ' Process Intervals Numbering Index
      n_step_interval = Worksheets("S3").Range("E13").Value
      n_interval = Worksheets("S3").Range("F13").Value
      For a = 1 To n_interval
         With Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1)
          .Value = n_step_interval
          .Interior.Color = RGB(221, 235, 247)
         End With
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(1)
          .Value = a
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter


'====================================================================
' [3c Part I] RAW MATERIAL SPECIFICATION: Selecting Basis Material
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(4)
      With head
         .Value = "[3] RAW MATERIAL MIXING SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      ' [BASIS MATERIAL] Sub-Section Title
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "Basis Material Specification"
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
    
      ' [BASIS MATERIAL] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [BASIS MATERIAL] List Components for Raw Materials Header
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Raw Material Addition- Basis Material Selection"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' [BASIS MATERIAL] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [BASIS MATERIAL] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter

' [3c Part II] RAW MATERIAL SPECIFICATION: Specific Raw Material Addition
      ' [RM ADDITION] Sub-Section Title
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "Raw Material Addition"
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
    
      ' [RM ADDITION] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [RM ADDITION] List Components for Raw Materials Header
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Specific Raw Material Addition (kg/kg-basis material)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' [RM ADDITION] Default RM Addition is 0 (no Raw Material Added)
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1, 2)
      With Range(cell_selected, cell_selected.Offset(n_proc_interval - 1, n_comp - 1))
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = 0
      End With
      
      ' [RM ADDITION] Energy Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp)
      With Range(cell_selected, cell_selected.Offset(, n_Eutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Energy Utilities: Specific Consumption (MJ/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(248, 203, 173)
      End With
      For a = 1 To n_Eutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp)
          .Value = Worksheets("B3").Range("C" & 4 + a).Value
          .Interior.Color = RGB(248, 203, 173)
         End With
      Next a

      ' [RM ADDITION] Mass Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp + n_Eutil)
      With Range(cell_selected, cell_selected.Offset(, n_Mutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Mass Utilities: Specific Consumption (kg/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(236, 134, 132)
      End With
      For a = 1 To n_Mutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp + n_Eutil)
          .Value = Worksheets("B4").Range("C" & 4 + a).Value
          .Interior.Color = RGB(236, 134, 132)
         End With
      Next a

      ' [RM ADDITION] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [RM ADDITION] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter


'====================================================================
' [3d Part I] REACTION SPECIFICATION: Key Reacting Component Fractional Conversion
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(4)
      With head
         .Value = "[4] REACTION SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With

      ' [KEY COMPONENT FRAC-CONV] Sub-Section Title
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "Key Component Fractional Conversion Specification"
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
      
      ' [KEY COMPONENT FRAC-CONV] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [KEY COMPONENT FRAC-CONV] Fractional Conversion for Key Reactant
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Key Reacting Component Fractional Conversion (0¡Â¥ö¡Â1)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' [KEY COMPONENT FRAC-CONV] Energy Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp)
      With Range(cell_selected, cell_selected.Offset(, n_Eutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Energy Utilities: Specific Consumption (MJ/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(248, 203, 173)
      End With
      For a = 1 To n_Eutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp)
          .Value = Worksheets("B3").Range("C" & 4 + a).Value
          .Interior.Color = RGB(248, 203, 173)
         End With
      Next a

      ' [KEY COMPONENT FRAC-CONV] Mass Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp + n_Eutil)
      With Range(cell_selected, cell_selected.Offset(, n_Mutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Mass Utilities: Specific Consumption (kg/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(236, 134, 132)
      End With
      For a = 1 To n_Mutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp + n_Eutil)
          .Value = Worksheets("B4").Range("C" & 4 + a).Value
          .Interior.Color = RGB(236, 134, 132)
         End With
      Next a

      ' [KEY COMPONENT FRAC-CONV] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [KEY COMPONENT FRAC-CONV] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter

' [3d Part II] REACTION SPECIFICATION: Non-Key Reacting Component Specific Consumption
      ' [NONKEY COMPONENT CONSUMPTION] Sub-Section Title
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "Non-Key Reacting Component Specific Consumption"
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
      
      ' [NONKEY COMPONENT CONSUMPTION] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [NONKEY COMPONENT CONSUMPTION] Specific Consumption of Non-key Reacting Components
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Non-Key Reacting Component Specific Consumption (kg/kg-key component reacted)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a

      ' [NONKEY COMPONENT CONSUMPTION] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [NONKEY COMPONENT CONSUMPTION] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter

' [3d Part III] REACTION SPECIFICATION: Product Fractional Yield
      ' [PRODUCT FRAC-YIELD] Sub-Section Title
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "Product Component Fractional Yield"
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
      
      ' [PRODUCT FRAC-YIELD] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [PRODUCT FRAC-YIELD] Fractional Yield of Products **NOTE** Row-Sum for each Interval Must Equal 1
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Fractional Yield of Product (¥Òi=1)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a

      ' [PRODUCT FRAC-YIELD] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [PRODUCT FRAC-YIELD] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter


'====================================================================
' [3e] WASTE PURGE SPECIFICATION
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "[5] WASTE PURGE SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      ' [WASTE PURGE] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [WASTE PURGE] Purge Fraction for Component i
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Purge Fraction for Component (0¡Â¥ä¡Â1)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a

      ' [WASTE PURGE] Default Waste Purge is 0 (aka no waste separated)
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1, 2)
      With Range(cell_selected, cell_selected.Offset(n_proc_interval - 1, n_comp - 1))
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = 0
      End With
      
      ' [WASTE PURGE] Energy Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp)
      With Range(cell_selected, cell_selected.Offset(, n_Eutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Energy Utilities: Specific Consumption (MJ/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(248, 203, 173)
      End With
      For a = 1 To n_Eutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp)
          .Value = Worksheets("B3").Range("C" & 4 + a).Value
          .Interior.Color = RGB(248, 203, 173)
         End With
      Next a

      ' [WASTE PURGE] Mass Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp + n_Eutil)
      With Range(cell_selected, cell_selected.Offset(, n_Mutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Mass Utilities: Specific Consumption (kg/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(236, 134, 132)
      End With
      For a = 1 To n_Mutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp + n_Eutil)
          .Value = Worksheets("B4").Range("C" & 4 + a).Value
          .Interior.Color = RGB(236, 134, 132)
         End With
      Next a

      ' [WASTE PURGE] Process Intervals Numbering Index
      For a = 1 To n_step - 2
         n_step_interval = Worksheets("S3").Range("E" & 13 + a).Value
         n_interval = Worksheets("S3").Range("F" & 13 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [WASTE PURGE] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter


'====================================================================
' [3f] PRODUCT SEPARATION SPECIFICATION
      ''=TITLE=
      Set head = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3)
      With head
         .Value = "[6] PRODUCT SEPARATION SPECIFICATION"
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      ' [PROD-SEPARATION] Process Intervals Table Index
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(2)
      With Range(cell_selected, cell_selected.Offset(, 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Index"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      With cell_selected.Offset(1)
         .Value = "Process Step"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(, 1)
      With cell_selected
         .Value = "Interval"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      
      ' [PROD-SEPARATION] Partition Fractions for Product Separation Into Primary Stream
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2)
      With Range(cell_selected, cell_selected.Offset(, n_comp - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Partition Fraction for Separation Into Primary Stream (0¡Â¥ò¡Â1)"
         .Font.Bold = True
         .Interior.Color = RGB(221, 235, 247)
      End With
      For a = 1 To n_comp
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a)
          .Value = Worksheets("B2").Range("C" & 3 + a).Value
          .Interior.Color = RGB(221, 235, 247)
         End With
      Next a
      
      ' [PROD-SEPARATION] Default product separation is 1 (aka, no separation, all streams go to primary)
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(1, 2)
      With Range(cell_selected, cell_selected.Offset(n_total_interval - 1, n_comp - 1))
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = 1
      End With
      
      ' [PROD-SEPARATION] Energy Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp)
      With Range(cell_selected, cell_selected.Offset(, n_Eutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Energy Utilities: Specific Consumption (MJ/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(248, 203, 173)
      End With
      For a = 1 To n_Eutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp)
          .Value = Worksheets("B3").Range("C" & 4 + a).Value
          .Interior.Color = RGB(248, 203, 173)
         End With
      Next a

      ' [PROD-SEPARATION] Mass Utilities Specification
      Set cell_selected = Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(-1, 2 + n_comp + n_Eutil)
      With Range(cell_selected, cell_selected.Offset(, n_Mutil - 1))
         .Merge
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .Value = "Mass Utilities: Specific Consumption (kg/kg-flow)"
         .Font.Bold = True
         .Interior.Color = RGB(236, 134, 132)
      End With
      For a = 1 To n_Mutil
         With Worksheets("B10").Cells(Rows.Count, "C").End(xlUp).Offset(, a + n_comp + n_Eutil)
          .Value = Worksheets("B4").Range("C" & 4 + a).Value
          .Interior.Color = RGB(236, 134, 132)
         End With
      Next a

      ' [PROD-SEPARATION] Process Intervals Numbering Index
      For a = 1 To n_step
         n_step_interval = Worksheets("S3").Range("E" & 12 + a).Value
         n_interval = Worksheets("S3").Range("F" & 12 + a).Value
         For b = 1 To n_interval
            With Cells(Rows.Count, "B").End(xlUp).Offset(1)
              .Value = n_step_interval
              .Interior.Color = RGB(221, 235, 247)
            End With
            With Cells(Rows.Count, "C").End(xlUp).Offset(1)
              .Value = b
              .Interior.Color = RGB(221, 235, 247)
            End With
         Next b
      Next a
      
      ' [PROD-SEPARATION] Drawing cell boundaries
      Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).CurrentRegion.Select
      Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
      Selection.HorizontalAlignment = xlCenter
      Selection.VerticalAlignment = xlCenter
      
      ' [Complete Checksum]
      Worksheets("B10").Cells(3, 5).Value = 1


'====================================================================
' [4] Create Named Ranges
    ' Declare Variables
    Dim IntervalNR As String
    Dim num_intervals As Integer
    
    ' Named Range for Process Intervals
    num_intervals = Worksheets("S3").Cells(14, 8).Value
    If Worksheets("B10").Cells(7, 2).Value <> 0 Then
        IntervalNR = "PI_IntervalNames"
        ActiveWorkbook.Names.Add Name:=IntervalNR, RefersToLocal:=Range(Cells(8, 2), Cells(7 + num_intervals, 4))
    End If
End Function



' [MASS BALANCE MODEL] Delete Interval Specification Table
Public Function TIPEM_Delete_IntervalSpecTable()
    ' [1] Define Variables
    Dim number_comp As Integer      'Compound number
    Dim number_n_Eutil As Integer   'Energy Utility number
    Dim number_n_Mutil As Integer   'Mass Utility number
    
    ' [2] Assign Values to Number Variables
    number_comp = Worksheets("B2").Range("K3").Value
    number_n_Eutil = Worksheets("B3").Range("C1").Value
    number_n_Mutil = Worksheets("B4").Range("C1").Value

    ' [3] Select Range to be Deleted to Default
    Worksheets("B10").Activate
    Worksheets("B10").Range(Worksheets("B10").Cells(4, 2), Worksheets("B10").Cells(Rows.Count, "B").End(xlUp).Offset(3, 6 + number_comp + number_n_Eutil + number_n_Mutil)).Select
    
    ' [4] Delete and Return to Default
    Selection.Font.Bold = False
    Selection.UnMerge
    Selection.ClearContents
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Interior.Pattern = xlNone
    Selection.Interior.TintAndShade = 0
    Selection.Interior.PatternTintAndShade = 0
    
    ' [Reset Checksum]
    Worksheets("B10").Cells(3, 5).Value = 0
End Function

