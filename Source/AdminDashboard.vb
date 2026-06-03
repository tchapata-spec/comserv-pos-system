Imports System.Data.OleDb

Public Class AdminDashboard
    Private dbConnection As String = GlobalVariables.DBConnectionString
    
    Private Sub AdminDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "ComServ Africa POS - Admin Dashboard"
        Me.BackColor = Color.WhiteSmoke
        Me.WindowState = FormWindowState.Maximized
        
        lblWelcome.Text = "Welcome, " & GlobalVariables.CurrentUsername & " (Admin)"
        LoadDashboardData()
    End Sub
    
    Private Sub LoadDashboardData()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                
                ' Load total products
                Dim cmdProducts As New OleDbCommand("SELECT COUNT(*) FROM tblProducts", conn)
                lblTotalProducts.Text = cmdProducts.ExecuteScalar().ToString()
                
                ' Load total sales
                Dim cmdSales As New OleDbCommand("SELECT COUNT(*) FROM tblSales", conn)
                lblTotalSales.Text = cmdSales.ExecuteScalar().ToString()
                
                ' Load total users
                Dim cmdUsers As New OleDbCommand("SELECT COUNT(*) FROM tblUsers", conn)
                lblTotalUsers.Text = cmdUsers.ExecuteScalar().ToString()
                
                ' Load low stock items
                Dim cmdLowStock As New OleDbCommand("SELECT ProductName, StockLevel FROM tblProducts WHERE StockLevel < 10 ORDER BY StockLevel ASC", conn)
                Using reader As OleDbDataReader = cmdLowStock.ExecuteReader()
                    lstLowStock.Items.Clear()
                    While reader.Read()
                        lstLowStock.Items.Add(reader("ProductName").ToString() & " - Stock: " & reader("StockLevel").ToString())
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading dashboard: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnManageProducts_Click(sender As Object, e As EventArgs) Handles btnManageProducts.Click
        Dim productForm As New ProductManagement()
        productForm.ShowDialog()
        LoadDashboardData()
    End Sub
    
    Private Sub btnManageUsers_Click(sender As Object, e As EventArgs) Handles btnManageUsers.Click
        Dim userForm As New UserManagement()
        userForm.ShowDialog()
        LoadDashboardData()
    End Sub
    
    Private Sub btnManagePromotions_Click(sender As Object, e As EventArgs) Handles btnManagePromotions.Click
        Dim promoForm As New PromotionManagement()
        promoForm.ShowDialog()
    End Sub
    
    Private Sub btnManageCustomers_Click(sender As Object, e As EventArgs) Handles btnManageCustomers.Click
        Dim custForm As New CustomerManagement()
        custForm.ShowDialog()
    End Sub
    
    Private Sub btnViewReports_Click(sender As Object, e As EventArgs) Handles btnViewReports.Click
        Dim reportForm As New SalesReports()
        reportForm.ShowDialog()
    End Sub
    
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If MessageBox.Show("Are you sure you want to logout?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Me.Close()
            Dim loginForm As New LoginForm()
            loginForm.Show()
        End If
    End Sub
    
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub
End Class