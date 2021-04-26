Public Class Form2
  Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
    PictureBox1.Location = New Point(0, 0)

  End Sub



  Private Sub Form2_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd
    PictureBox1.Size = Me.Size
  End Sub
End Class