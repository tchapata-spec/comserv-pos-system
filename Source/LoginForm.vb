Imports System.Data.OleDb

Public Class LoginForm
    Private dbConnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\ComServPOS.accdb"
    
    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "ComServ Africa POS - Login"
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen
        
        ' Set focus to username
        txtUsername.Focus()
    End Sub
    
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If ValidateLogin() Then
            AuthenticateUser()
        End If
    End Sub
    
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub
    
    Private Function ValidateLogin() As Boolean
        If String.IsNullOrWhiteSpace(txtUsername.Text) Then
            MessageBox.Show("Please enter username", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsername.Focus()
            Return False
        End If
        
        If String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("Please enter password", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPassword.Focus()
            Return False
        End If
        
        Return True
    End Function
    
    Private Sub AuthenticateUser()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                
                Dim query As String = "SELECT UserID, Username, Role FROM tblUsers WHERE Username = @username AND Password = @password AND IsActive = True"
                
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@username", txtUsername.Text)
                    cmd.Parameters.AddWithValue("@password", txtPassword.Text)
                    
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        If reader.HasRows Then
                            reader.Read()
                            Dim userID As Integer = CInt(reader("UserID"))
                            Dim username As String = reader("Username").ToString()
                            Dim role As String = reader("Role").ToString()
                            
                            ' Store user info in global variables
                            GlobalVariables.CurrentUserID = userID
                            GlobalVariables.CurrentUsername = username
                            GlobalVariables.CurrentUserRole = role
                            
                            ' Open appropriate form based on role
                            If role = "Admin" Then
                                Dim adminForm As New AdminDashboard()
                                adminForm.Show()
                            ElseIf role = "Teller" Then
                                Dim tellerForm As New TellerDashboard()
                                tellerForm.Show()
                            End If
                            
                            Me.Hide()
                        Else
                            MessageBox.Show("Invalid username or password", "Authentication Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            txtPassword.Clear()
                            txtUsername.Focus()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = Chr(13) Then ' Enter key
            btnLogin_Click(Nothing, Nothing)
            e.Handled = True
        End If
    End Sub
    
    Private Sub txtUsername_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUsername.KeyPress
        If e.KeyChar = Chr(13) Then ' Enter key
            txtPassword.Focus()
            e.Handled = True
        End If
    End Sub
End Class