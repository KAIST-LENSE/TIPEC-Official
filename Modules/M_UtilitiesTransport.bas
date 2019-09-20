Attribute VB_Name = "M_UtilitiesTransport"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''         TIPEM U&T MODULE        ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ¿Ã¡ˆ»Ø/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [UTILITIES & TRANSPORT] Mouse Over with Display Table for Energy Utility
Public Function M_MouseOverEvent_EnergyUtility()
    ' [1] Change UI (none required for Energy Utility)
    Sheets("S2").Range("G13").Value = "Index"
    Sheets("S2").Range("H13").Value = "Utility Name"
    Sheets("S2").Range("J13").Value = "CO2 Footprint (ton CO2e/GJ)"
    Sheets("S2").Range("L14").Value = "($/GJ)"

    ' [2] Change Display Table Entries
    ' Energy Utility Index
    Set UtilityE_Index_UpTo20 = Sheets("B3").Range("B5:B24")
    Set DisplayE_Index_UpTo20 = Sheets("S2").Range("G15:G34")
    DisplayE_Index_UpTo20.Value = UtilityE_Index_UpTo20.Value
    ' Energy Utility Name
    Set UtilityE_Name_UpTo20 = Sheets("B3").Range("C5:C24")
    Set DisplayE_Name_UpTo20 = Sheets("S2").Range("H15:H34")
    DisplayE_Name_UpTo20.Value = UtilityE_Name_UpTo20.Value
    ' Energy Utility CO2 Footprint Production
    Set UtilityE_CO2P_UpTo20 = Sheets("B3").Range("D5:D24")
    Set DisplayE_CO2P_UpTo20 = Sheets("S2").Range("J15:J34")
    DisplayE_CO2P_UpTo20.Value = UtilityE_CO2P_UpTo20.Value
    ' Energy Utility CO2 Footprint Consumption
    Set UtilityE_CO2C_UpTo20 = Sheets("B3").Range("E5:E24")
    Set DisplayE_CO2C_UpTo20 = Sheets("S2").Range("K15:K34")
    DisplayE_CO2C_UpTo20.Value = UtilityE_CO2C_UpTo20.Value
    ' Energy Utility Specific Cost
    Set UtilityE_Cost_UpTo20 = Sheets("B3").Range("F5:F24")
    Set DisplayE_Cost_UpTo20 = Sheets("S2").Range("L15:L34")
    DisplayE_Cost_UpTo20.Value = UtilityE_Cost_UpTo20.Value
    
    ' [3] Change Display Table Color
    ' Fill Color
    Range("G13:L34").Interior.Color = RGB(221, 235, 247)
    
    ' Tab Disable Borders
    With Range("G11:I12")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(221, 235, 247)
    End With
    
    ' Tab Enable Other Border, add bottom edge
    With Range("J11:L12")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    End With

End Function



' [UTILITIES & TRANSPORT] Mouse Over with Display Table for Mass Utility
Public Function M_MouseOverEvent_MassUtility()
    ' [1] Change UI (none required for Energy Utility)
    Sheets("S2").Range("G13").Value = "Index"
    Sheets("S2").Range("H13").Value = "Utility Name"
    Sheets("S2").Range("J13").Value = "CO2 Footprint (ton CO2e/ton)"
    Sheets("S2").Range("L14").Value = "($/ton)"

    ' [2] Change Display Table Entries
    ' Mass Utility Index
    Set UtilityM_Index_UpTo20 = Sheets("B4").Range("B5:B24")
    Set DisplayM_Index_UpTo20 = Sheets("S2").Range("G15:G34")
    DisplayM_Index_UpTo20.Value = UtilityM_Index_UpTo20.Value
    ' Mass Utility Name
    Set UtilityM_Name_UpTo20 = Sheets("B4").Range("C5:C24")
    Set DisplayM_Name_UpTo20 = Sheets("S2").Range("H15:H34")
    DisplayM_Name_UpTo20.Value = UtilityM_Name_UpTo20.Value
    ' Mass Utility CO2 Footprint Production
    Set UtilityM_CO2P_UpTo20 = Sheets("B4").Range("D5:D24")
    Set DisplayM_CO2P_UpTo20 = Sheets("S2").Range("J15:J34")
    DisplayM_CO2P_UpTo20.Value = UtilityM_CO2P_UpTo20.Value
    ' Mass Utility CO2 Footprint Consumption
    Set UtilityM_CO2C_UpTo20 = Sheets("B4").Range("E5:E24")
    Set DisplayM_CO2C_UpTo20 = Sheets("S2").Range("K15:K34")
    DisplayM_CO2C_UpTo20.Value = UtilityM_CO2C_UpTo20.Value
    ' Mass Utility Specific Cost
    Set UtilityM_Cost_UpTo20 = Sheets("B4").Range("F5:F24")
    Set DisplayM_Cost_UpTo20 = Sheets("S2").Range("L15:L34")
    DisplayM_Cost_UpTo20.Value = UtilityM_Cost_UpTo20.Value
    
    
    ' [3] Change Display Table Color
    ' Fill Color
    Range("G13:L34").Interior.Color = RGB(248, 203, 173)
    
    ' Tab Disable Borders
    With Range("J11:L12")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(248, 203, 173)
    End With
    
    ' Tab Enable Other Border, add bottom edge
    With Range("G11:I12")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    End With
End Function

