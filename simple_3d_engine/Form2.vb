Imports System.ComponentModel

Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.BringToFront()
        TextBox2.SelectionLength = 0
    End Sub

    Private Sub Form2_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Form1.Enabled = True
    End Sub
End Class