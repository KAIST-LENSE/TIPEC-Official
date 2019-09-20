VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' [1] Update Materials List Table when Scrollbar is Activated
Private Sub ScrollBar2_Change()
    S1.Range("E13:M32").Value = S3_2.Range(CStr(Mid(S1.Range("E32").Value, 2, 1)) + CStr(ScrollBar2.Value) + ":" + CStr(Mid(S1.Range("M32").Value, 2, 1)) + CStr(ScrollBar2.Value + 20)).Value
End Sub

