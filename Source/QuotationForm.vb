Imports System.Data.OleDb

Public Class QuotationForm
    Private dbConnection As String = GlobalVariables.DBConnectionString
    Private currentQuotationItems As New List(Of QuotationItem)
    
    Private Class QuotationItem
        Public ProductID As Integer
        Public ProductName As String
        Public Quantity As Integer
        Public UnitPrice As Decimal
        Public LineTotal As Decimal
    End Class
    
    Private Sub QuotationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Quotation Management"
        LoadProducts()
    End Sub
    
    Private Sub LoadProducts()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT ProductID, ProductName, UnitPrice FROM tblProducts ORDER BY ProductName", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                dgvAvailableProducts.DataSource = table
                dgvAvailableProducts.AutoResizeColumns()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading products: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
