VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Specify Distribution for Process Variable"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()
If ComboBox1.Value = "Triangular" Then
    Me.Frame1.Visible = True
    Me.Frame2.Visible = False
    Me.Frame3.Visible = False
ElseIf ComboBox1.Value = "Normal" Then
    Me.Frame1.Visible = False
    Me.Frame2.Visible = True
    Me.Frame3.Visible = False
ElseIf ComboBox1.Value = "Uniform" Then
    Me.Frame1.Visible = False
    Me.Frame2.Visible = False
    Me.Frame3.Visible = True
End If
End Sub
