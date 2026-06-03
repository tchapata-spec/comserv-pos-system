Imports System.Data.OleDb

Public Class ReturnsForm
    Private dbConnection As String = GlobalVariables.DBConnectionString
    
    Private Sub ReturnsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Returns Management"
        LoadSales()
    End Sub
    
    Private Sub LoadSales()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT SaleID, TellerName, SaleDate, TotalAmount FROM tblSales ORDER BY SaleDate DESC", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                dgvSales.DataSource = table
                dgvSales.AutoResizeColumns()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading sales: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
