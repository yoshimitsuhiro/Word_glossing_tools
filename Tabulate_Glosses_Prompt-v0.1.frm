VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tabulate_Glosses_Prompt 
   Caption         =   "UserForm1"
   ClientHeight    =   3072
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   4608
   OleObjectBlob   =   "Tabulate_Glosses_Prompt-v0.1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tabulate_Glosses_Prompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private userCancelled As Boolean

Private Sub btnOK_Click()
    If (TextBox1.Value <> "Auto" And Not IsNumeric(TextBox1.Value)) Or Not IsNumeric(TextBox2.Value) Then
        MsgBox "Please enter ""Auto"" or a number in the first box, and a number in the second box.", vbExclamation
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    userCancelled = True
    Me.Hide
End Sub

Public Function ShowForm(ByRef indentStr As String, ByRef interval As Single) As Boolean
    Me.TextBox1.Value = "Auto"
    Me.TextBox2.Value = "10"
    userCancelled = False
    Me.Show
    If userCancelled Then
        ShowForm = False
    Else
        If Me.TextBox1.Value = "Auto" Then
            indentStr = "Auto"
        Else
            indentStr = CStr(Me.TextBox1.Value)
        End If
        interval = CSng(Me.TextBox2.Value)
        ShowForm = True
    End If
End Function

Private Sub UserForm_Click()

End Sub
