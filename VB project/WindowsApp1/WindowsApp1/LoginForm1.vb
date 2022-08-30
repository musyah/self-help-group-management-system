Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim a As String
        Dim b As String
        Dim password1 = "Michael123"
        Dim user1 = "Michael"
        Dim password2 = "Musya123"
        Dim user2 = "Musya"

        a = PasswordTextBox.Text
        b = UsernameTextBox.Text

        If (a = password1 And b = user1) Or (a = password2 And b = user2) Then
            Home.Show()
            MessageBox.Show("Successful Log In")
        Else
            MessageBox.Show("Incorrect username or password")
        End If
        PasswordTextBox.Text = ""
        UsernameTextBox.Text = ""
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub PasswordLabel_Click(sender As Object, e As EventArgs) Handles PasswordLabel.Click

    End Sub
End Class
