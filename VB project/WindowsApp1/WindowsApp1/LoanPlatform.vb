Imports System.Data.SqlClient

Public Class LoanPlatform
    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim i As Integer
    Private Sub LoanPlatform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Const V As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\user\Documents\Final year project\WindowsApp1\WindowsApp1\Main.mdf;Integrated Security=True"
        con.ConnectionString = V
        If con.State = ConnectionState.Open Then
            con.Close()

        End If
        con.Open()

        disp_data()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "insert into Loans values('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox9.Text + "','" + TextBox5.Text + "','" + TextBox6.Text + "','" + TextBox7.Text + "','" + TextBox10.Text + "','" + TextBox8.Text + "')"
        cmd.ExecuteNonQuery()

        disp_data()
        TextBox3.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""



        MessageBox.Show("Record Added Successfully")

    End Sub
    Public Sub disp_data()

        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select * from Loans"
        cmd.ExecuteNonQuery()
        Dim dt As New DataTable()
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(dt)
        DataGridView1.DataSource = dt

    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick


        Try
            If con.State = ConnectionState.Open Then
                con.Close()

            End If
            con.Open()

            i = Convert.ToInt32(DataGridView1.SelectedCells.Item(0).Value.ToString())

            cmd = con.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "select * from Loans where id=" & i & ""
            cmd.ExecuteNonQuery()
            Dim dt As New DataTable()
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(dt)
            Dim dr As SqlClient.SqlDataReader
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            While dr.Read

                TextBox1.Text = dr.GetString(1).ToString()
                TextBox2.Text = dr.GetString(2).ToString()

            End While
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If con.State = ConnectionState.Open Then
            con.Close()

        End If
        con.Open()

        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "delete from Loans where LID='" + TextBox1.Text + "'"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Delete Successful")

        disp_data()
        TextBox1.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select * from Loans where LID='" + TextBox11.Text + "'"
        cmd.ExecuteNonQuery()
        Dim dt As New DataTable()
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(dt)
        DataGridView1.DataSource = dt

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        disp_data()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim z As Integer
        Dim w As Integer
        Dim d As Integer


        z = TextBox4.Text
        w = TextBox9.Text
        d = z * 4 * w
        TextBox5.Text = d
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim c As Integer
        Dim a As Integer
        Dim b As Integer
        Dim d As Integer


        a = TextBox6.Text
        b = TextBox7.Text

        If (b < 6) Then
            d = (a + (a * 0.06))
            c = d / b

        Else
            d = (a + (a * 0.045))
            c = d / b
        End If

        TextBox8.Text = c
        TextBox10.Text = d



    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim a As Integer
        Dim b As Integer
        Dim p As Integer
        Dim c As Integer

        b = TextBox6.Text
        a = TextBox5.Text
        p = TextBox7.Text
        c= TextBox9.Text

        If b > a Then
            MessageBox.Show("Amount borrowed exceeds Amount Qualified")
        ElseIf p > 15 Then
            MessageBox.Show("Repayment Period cannot exceed 15 months")
        ElseIf c < 1.2 Then
            MessageBox.Show("You cannot Qualify for a Loan due to low credit score")
        Else
            MessageBox.Show("Loan documentation approved. Wait for approval and disbursement")
        End If



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Contributions.Show()
        Repayments.Close()
        Me.Close()

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Contributions.Close()
        Repayments.Close()
        Registry.Close()
        Me.Close()
        Home.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Repayments.Show()

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub
End Class