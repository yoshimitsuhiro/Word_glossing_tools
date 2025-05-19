VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tabulate_Glosses_Results 
   Caption         =   "Results"
   ClientHeight    =   5488
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   7808
   OleObjectBlob   =   "Tabulate_Glosses_Results-v0.2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tabulate_Glosses_Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowText(ByVal output As String)

    Me.DebugTextBox.Text = output
    Me.Show

End Sub

Private Sub DebugTextBox_Change()

End Sub
