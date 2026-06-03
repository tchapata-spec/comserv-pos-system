Imports System.Data.OleDb

Public Class ProductManagement
    Private dbConnection As String = GlobalVariables.DBConnectionString
    Private editingProductID As Integer = -1
    
    Private Sub ProductManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Product Management"
        LoadProducts()
    End Sub
    
    Private Sub LoadProducts()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                
                Dim cmd As New OleDbCommand("SELECT ProductID, ProductName, Specs, UnitPrice, StockLevel FROM tblProducts ORDER BY ProductName", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                
                dgvProducts.DataSource = table
                dgvProducts.AutoResizeColumns()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading products: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnAddProduct_Click(sender As Object, e As EventArgs) Handles btnAddProduct.Click
        If ValidateProductInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("INSERT INTO tblProducts (ProductName, Specs, UnitPrice, StockLevel) VALUES (@name, @specs, @price, @stock)", conn)
                    cmd.Parameters.AddWithValue("@name", txtProductName.Text)
                    cmd.Parameters.AddWithValue("@specs", txtSpecs.Text)
                    cmd.Parameters.AddWithValue("@price", CDec(txtUnitPrice.Text))
                    cmd.Parameters.AddWithValue("@stock", CInt(numStockLevel.Value))
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Product added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearProductForm()
                    LoadProducts()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error adding product: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub btnUpdateProduct_Click(sender As Object, e As EventArgs) Handles btnUpdateProduct.Click
        If editingProductID = -1 Then
            MessageBox.Show("Please select a product to update", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        
        If ValidateProductInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("UPDATE tblProducts SET ProductName = @name, Specs = @specs, UnitPrice = @price, StockLevel = @stock, LastModified = @modified WHERE ProductID = @id", conn)
                    cmd.Parameters.AddWithValue("@name", txtProductName.Text)
                    cmd.Parameters.AddWithValue("@specs", txtSpecs.Text)
                    cmd.Parameters.AddWithValue("@price", CDec(txtUnitPrice.Text))
                    cmd.Parameters.AddWithValue("@stock", CInt(numStockLevel.Value))
                    cmd.Parameters.AddWithValue("@modified", DateTime.Now)
                    cmd.Parameters.AddWithValue("@id", editingProductID)
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Product updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearProductForm()
                    editingProductID = -1
                    LoadProducts()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error updating product: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub dgvProducts_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProducts.CellDoubleClick
        If e.RowIndex >= 0 Then
            editingProductID = CInt(dgvProducts.Rows(e.RowIndex).Cells("ProductID").Value)
            txtProductName.Text = dgvProducts.Rows(e.RowIndex).Cells("ProductName").Value.ToString()
            txtSpecs.Text = dgvProducts.Rows(e.RowIndex).Cells("Specs").Value.ToString()
            txtUnitPrice.Text = dgvProducts.Rows(e.RowIndex).Cells("UnitPrice").Value.ToString()
            numStockLevel.Value = CInt(dgvProducts.Rows(e.RowIndex).Cells("StockLevel").Value)
        End If
    End Sub
    
    Private Function ValidateProductInput() As Boolean
        If String.IsNullOrWhiteSpace(txtProductName.Text) Then
            MessageBox.Show("Please enter product name", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductName.Focus()
            Return False
        End If
        
        If Not Decimal.TryParse(txtUnitPrice.Text, Nothing) Then
            MessageBox.Show("Please enter a valid unit price", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUnitPrice.Focus()
            Return False
        End If
        
        If CDec(txtUnitPrice.Text) <= 0 Then
            MessageBox.Show("Unit price must be greater than 0", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUnitPrice.Focus()
            Return False
        End If
        
        Return True
    End Function
    
    Private Sub ClearProductForm()
        txtProductName.Clear()
        txtSpecs.Clear()
        txtUnitPrice.Clear()
        numStockLevel.Value = 0
        editingProductID = -1
    End Sub
    
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearProductForm()
    End Sub
    
    Private Sub btnDeleteProduct_Click(sender As Object, e As EventArgs) Handles btnDeleteProduct.Click
        If dgvProducts.SelectedRows.Count > 0 Then
            If MessageBox.Show("Are you sure you want to delete this product?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    Dim productID As Integer = CInt(dgvProducts.SelectedRows(0).Cells("ProductID").Value)
                    
                    Using conn As New OleDbConnection(dbConnection)
                        conn.Open()
                        
                        Dim cmd As New OleDbCommand("DELETE FROM tblProducts WHERE ProductID = @id", conn)
                        cmd.Parameters.AddWithValue("@id", productID)
                        
                        cmd.ExecuteNonQuery()
                        
                        MessageBox.Show("Product deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ClearProductForm()
                        LoadProducts()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error deleting product: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a product to delete", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class