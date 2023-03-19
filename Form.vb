Public Class Form
    Private Sub Form_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim userName = Val(TextBox1.Text)
        Dim userPassword = Val(TextBox2.Text)
        If (userName = "Dianah" And
            userPassword = "123") Then
            Form1.Show()

        End If
    End Sub
End Class