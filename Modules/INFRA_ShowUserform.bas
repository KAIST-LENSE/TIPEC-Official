Attribute VB_Name = "INFRA_ShowUserform"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''    TIPEM USERFORM SHOW MODULE   ''''''''''
''''''''                                 ''''''''''
'''''''' Made by ¿Ã¡ˆ»Ø/Steve Lee, 2018  ''''''''''
''''''''          KAIST, LENSE           ''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
' [INFRA] Show UF Create Project
Sub S0_CreateProject()
U0a_CreateProject.Show
End Sub


' [INFRA] Show UF Add a Material
Sub S1_AddaMaterial()
U2a_MaterialAdd.Show
End Sub


' [INFRA] Show UF Edit/Remove a Material
Sub S1_EditorRemoveMaterial()
U2b_MaterialEditRemove.Show
End Sub


' [INFRA] Show UF Add a Utility
Sub S2_AddaUtility()
U3a_UtilityAdd.Show
End Sub


' [INFRA] Show UF Add a Transport
Sub S2_AddaTransportation()
U3e_TransportAdd.Show
End Sub


' [INFRA] Show UF Edit/Remove a Utility
Sub S2_EditRemoveUtility()
U3b_UtilityEditRemove.Show
End Sub


' [INFRA] Show UF Edit/Remove a Transport
Sub S2_EditRemoveTransport()
U3f_TransportEditRemove.Show
End Sub


' [INFRA] Assign Custom Interval Names
Sub S4_AssignIntervalNames()
U5g_AssignIntervalNames.Show
End Sub


' [INFRA] Transportation Distance
Sub S4_TransportDistanceSpec()
U5j_TransportDistance.Show
End Sub


' [INFRA] Specify Feedstock Composition and Feedrate
Sub S4_FeedstockSpec()
On Error Resume Next
U5h_Feedstock_Specification.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: INPUT STREAMS
Sub S4_InputStreams()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Oval 58")).ZOrder msoSendToFront
    
    ' Show Userform
    U5a_StreamsIn.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: MIXING
Sub S4_Mixing()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Oval 59")).ZOrder msoSendToFront
    
    ' Show Userform
    U5b_Mixing.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: REACTION
Sub S4_Reaction()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Group 60")).ZOrder msoSendToFront

    ' Show Userform
    U5c_Reaction.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: WASTE PURGE
Sub S4_WastePurge()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Diamond 64")).ZOrder msoSendToFront
    
    ' Show Userform
    U5d_WastePurge.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: SEPARATION
Sub S4_Separation()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Flowchart: Sort 65")).ZOrder msoSendToFront

    ' Show Userform
    U5e_Separation.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: OUTPUT STREAMS 1
Sub S4_OutputStreams1()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Oval 66")).ZOrder msoSendToFront
    
    ' Show Userform
    U5f_StreamsOut.Show
End Sub


' [INFRA] Show UF PROCESS SPEC: OUTPUT STREAMS 2
Sub S4_OutputStreams2()
    ' Select Figure as Button
    Application.ScreenUpdating = True
    ActiveSheet.Shapes.Range(Array("Oval 67")).ZOrder msoSendToFront
    
    ' Show Userform
    U5f_StreamsOut.Show
End Sub


' [INFRA] Show TEA: Equipment Cost Userform
Sub S5_TEA_Equipment()
    ' Show Userform
    U7a_EquipmentCost.Show
End Sub


' [INFRA] Show TEA: Lang Factors Userform
Sub S5_TEA_LangFactors()
    ' Show Userform
    U7b_LangFactors.Show
End Sub


' [INFRA] Show TEA: TEA Params Userform
Sub S5_TEA_Params()
    ' Show Userform
    U7c_TEA_Parameters.Show
End Sub


' [INFRA] Conduct DCFROR
Sub S5_DCFROR()
    ' Show Userform
    U7d_DCFROR.Show
End Sub


' [INFRA] Conduct TEA Evaluation
Sub S7_Evaluate_TEA()
    ' Show Userform
    U7e_Evaluate_TEA.Show
End Sub


' [INFRA] Show Userform to Select from Available Connections
Sub S8_Connections_Choose()
    ' Show Userform
    U6a_Pathway_Specification.Show
End Sub
