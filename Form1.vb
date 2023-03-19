Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Data.OleDb



Public Class Form1
    Dim conn As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim dt As New DataTable
    Dim da As New OleDbDataAdapter(cmd)

    Dim num1 As Double
    Dim num2 As Double
    Dim result As Double
    Dim operation As String
    Dim TotalCost As Double
    Dim TotalDeliveryCost As Double
    Dim deliveryprice As Double
    Dim TotalMileageCost As Double
    Dim fabrics = 2000
    Dim blinds = 3000
    Dim carpet = 1000
    Dim mileagePrice = 88






    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        num1 = Val(TextBox1.Text)
        operation = "-"
        Label1.Text += "-"
        TextBox1.Text = ""
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        num1 = Val(TextBox1.Text)
        operation = "+"
        Label1.Text += "+"
        TextBox1.Text = ""
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        num1 = Val(TextBox1.Text)
        operation = "-"
        Label1.Text += "-"
        TextBox1.Text = ""
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        num1 = Val(TextBox1.Text)
        operation = "/"
        Label1.Text += "/"
        TextBox1.Text = ""
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        num1 = Val(TextBox1.Text)
        operation = "*"
        Label1.Text += "*"
        TextBox1.Text = ""
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        ' Equals operation
        num2 = Val(TextBox1.Text)

        Select Case operation
            Case "+"
                result = num1 + num2
            Case "-"
                result = num1 - num2
            Case "*"
                result = num1 * num2
            Case "/"
                result = num1 / num2
        End Select
        Label1.Text = result.ToString()
        TextBox1.Text = ""
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text += "1"
        TextBox1.Text += "1"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label1.Text += "2"
        TextBox1.Text += "2"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Button for number 3
        Label1.Text += "3"
        TextBox1.Text += "3"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Button for number 4
        Label1.Text += "4"
        TextBox1.Text += "4"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Button for number 5
        Label1.Text += "5"
        TextBox1.Text += "5"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Button for number 6
        Label1.Text += "6"
        TextBox1.Text += "6"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' Button for number 7
        Label1.Text += "7"
        TextBox1.Text += "7"
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Button for number 8
        Label1.Text += "8"
        TextBox1.Text += "8"
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Button for number 9
        TextBox1.Text += "9"
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        If Not (TextBox1.Text.Contains("."c)
        ) Then
            Label1.Text += "."
            TextBox1.Text += "."
        End If
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ' Button for number 9
        Label1.ResetText()
        TextBox1.Clear()

    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\deand\Documents\OSS.accdb"



    End Sub



    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TotalCost = ((carpet * Val(TextBox3.Text)) + (fabrics * Val(TextBox4.Text)) + (blinds * Val(TextBox5.Text)))
        TextBox12.Text = TotalCost.ToString() + "RWF"
        deliveryprice = Val(TextBox6.Text)
        TextBox6.Text = deliveryprice.ToString() + "RWF"
        TotalDeliveryCost = ((deliveryprice * Val(TextBox3.Text)) + (deliveryprice * Val(TextBox4.Text)) + (deliveryprice * Val(TextBox5.Text)))
        TextBox8.Text = TotalDeliveryCost.ToString() + "RWF"
        TotalMileageCost = (Val(TextBox7.Text) * mileagePrice)
        TextBox9.Text = TotalMileageCost.ToString() + "RWF"
        TextBox13.Text = (TotalCost + TotalDeliveryCost + TotalMileageCost).ToString() + "RWF"
        TextBox11.Text = "18%"
        TextBox10.Text = ((TotalCost + TotalDeliveryCost + TotalMileageCost) + Val(TextBox13.Text) * 0.18).ToString()
        If (TextBox2.Text = "") Then
            TextBox2.AppendText(" 3D ONLINE SHOPPING SYSTEM" + vbNewLine)
            TextBox2.AppendText("==================" + vbNewLine)
            TextBox2.AppendText(" Carpets ==" + Val(TextBox3.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Fabrics ==" + Val(TextBox4.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Blinds ==" + Val(TextBox5.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Delivery Fee ==" + Val(TextBox8.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Shipping Cost  ==" + Val(TextBox9.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Tax ==" + Val(TextBox11.Text).ToString() + vbNewLine)
            TextBox2.AppendText(" Total Cost ==" + Val(TextBox10.Text).ToString() + "RWF" + vbNewLine)

        End If



    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox13.Text = ""

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        conn.Open()
        cmd = conn.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "insert into oss(Carpets,Fabrics,Blinds,DeliveryFee,ShippingCost,Tax,TotalCost) values('" + TextBox3.Text + "','" + TextBox4.Text + "', '" + TextBox5.Text + "','" + TextBox8.Text + "','" + TextBox9.Text + "','" + TextBox11.Text + "','" + TextBox10.Text + "')"
        cmd.ExecuteNonQuery()
        conn.Close()

    End Sub
End Class
