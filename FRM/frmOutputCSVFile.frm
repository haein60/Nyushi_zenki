VERSION 5.00
Begin VB.Form frmOutputCSVFile 
   Caption         =   "1096"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmOutputCSVFile.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   8580
   WindowState     =   2  '�ő剻
End
Attribute VB_Name = "frmOutputCSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    fMainForm.mnuTools.Enabled = False  ' disable tools menu
    Dim Index
    For Index = 1 To fMainForm.Toolbar1.Buttons.Count
       fMainForm.Toolbar1.Buttons(Index).Enabled = False
    Next
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, LoadResString(1729)
End Sub

