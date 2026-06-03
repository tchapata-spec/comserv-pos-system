Imports System.Data.OleDb

Public Class CustomerManagement
    Private dbConnection As String = GlobalVariables.DBConnectionString
    Private editingCustomerID As Integer = -1
    
    Private Sub CustomerManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Customer Management"
        LoadCustomers()
    End Sub
    
    Private Sub LoadCustomers()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                
                Dim cmd As New OleDbCommand("SELECT CustomerID, CustomerName, ContactNumber, EmailAddress, Address FROM tblCustomers ORDER BY CustomerName", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                
                dgvCustomers.DataSource = table
                dgvCustomers.AutoResizeColumns()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading customers: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnAddCustomer_Click(sender As Object, e As EventArgs) Handles btnAddCustomer.Click
        If ValidateCustomerInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("INSERT INTO tblCustomers (CustomerName, ContactNumber, EmailAddress, Address) VALUES (@name, @contact, @email, @address)", conn)
                    cmd.Parameters.AddWithValue("@name", txtCustomerName.Text)
                    cmd.Parameters.AddWithValue("@contact", txtContactNumber.Text)
                    cmd.Parameters.AddWithValue("@email", txtEmail.Text)
                    cmd.Parameters.AddWithValue("@address", txtAddress.Text)
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Customer added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearCustomerForm()
                    LoadCustomers()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error adding customer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub btnUpdateCustomer_Click(sender As Object, e As EventArgs) Handles btnUpdateCustomer.Click
        If editingCustomerID = -1 Then
            MessageBox.Show("Please select a customer to update", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        
        If ValidateCustomerInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("UPDATE tblCustomers SET CustomerName = @name, ContactNumber = @contact, EmailAddress = @email, Address = @address WHERE CustomerID = @id", conn)
                    cmd.Parameters.AddWithValue("@name", txtCustomerName.Text)
                    cmd.Parameters.AddWithValue("@contact", txtContactNumber.Text)
                    cmd.Parameters.AddWithValue("@email", txtEmail.Text)
                    cmd.Parameters.AddWithValue("@address", txtAddress.Text)
                    cmd.Parameters.AddWithValue("@id", editingCustomerID)
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Customer updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearCustomerForm()
                    editingCustomerID = -1
                    LoadCustomers()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error updating customer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub btnBlacklistCustomer_Click(sender As Object, e As EventArgs) Handles btnBlacklistCustomer.Click
        If dgvCustomers.SelectedRows.Count > 0 Then
            Dim customerID As Integer = CInt(dgvCustomers.SelectedRows(0).Cells("CustomerID").Value)
            Dim customerName As String = dgvCustomers.SelectedRows(0).Cells("CustomerName").Value.ToString()
            
            Dim reason As String = InputBox("Enter reason for blacklisting:", "Blacklist Customer")
            
            If Not String.IsNullOrWhiteSpace(reason) Then
                Try
                    Using conn As New OleDbConnection(dbConnection)
                        conn.Open()
                        
                        Dim cmd As New OleDbCommand("INSERT INTO tblBlacklist (CustomerID, Reason, AddedBy) VALUES (@customerID, @reason, @addedBy)", conn)
                        cmd.Parameters.AddWithValue("@customerID", customerID)
                        cmd.Parameters.AddWithValue("@reason", reason)
                        cmd.Parameters.AddWithValue("@addedBy", GlobalVariables.CurrentUsername)
                        
                        cmd.ExecuteNonQuery()
                        
                        MessageBox.Show(customerName & " has been blacklisted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LoadCustomers()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error blacklisting customer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a customer to blacklist", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnWhitelistCustomer_Click(sender As Object, e As EventArgs) Handles btnWhitelistCustomer.Click
        If dgvCustomers.SelectedRows.Count > 0 Then
            Dim customerID As Integer = CInt(dgvCustomers.SelectedRows(0).Cells("CustomerID").Value)
            Dim customerName As String = dgvCustomers.SelectedRows(0).Cells("CustomerName").Value.ToString()
            
            Dim discountStr As String = InputBox("Enter discount percentage for whitelisting (0-100):", "Whitelist Customer", "0")
            
            If Not String.IsNullOrWhiteSpace(discountStr) AndAlso Decimal.TryParse(discountStr, Nothing) Then
                Try
                    Using conn As New OleDbConnection(dbConnection)
                        conn.Open()
                        
                        Dim cmd As New OleDbCommand("INSERT INTO tblWhitelist (CustomerID, DiscountPercentage, ValidFrom, ValidTo) VALUES (@customerID, @discount, @validFrom, @validTo)", conn)
                        cmd.Parameters.AddWithValue("@customerID", customerID)
                        cmd.Parameters.AddWithValue("@discount", CDec(discountStr))
                        cmd.Parameters.AddWithValue("@validFrom", DateTime.Now)
                        cmd.Parameters.AddWithValue("@validTo", DateTime.Now.AddYears(1))
                        
                        cmd.ExecuteNonQuery()
                        
                        MessageBox.Show(customerName & " has been whitelisted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LoadCustomers()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error whitelisting customer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                MessageBox.Show("Please enter a valid discount percentage", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            MessageBox.Show("Please select a customer to whitelist", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub dgvCustomers_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCustomers.CellDoubleClick
        If e.RowIndex >= 0 Then
            editingCustomerID = CInt(dgvCustomers.Rows(e.RowIndex).Cells("CustomerID").Value)
            txtCustomerName.Text = dgvCustomers.Rows(e.RowIndex).Cells("CustomerName").Value.ToString()
            txtContactNumber.Text = dgvCustomers.Rows(e.RowIndex).Cells("ContactNumber").Value.ToString()
            txtEmail.Text = dgvCustomers.Rows(e.RowIndex).Cells("EmailAddress").Value.ToString()
            txtAddress.Text = dgvCustomers.Rows(e.RowIndex).Cells("Address").Value.ToString()
        End If
    End Sub
    
    Private Function ValidateCustomerInput() As Boolean
        If String.IsNullOrWhiteSpace(txtCustomerName.Text) Then
            MessageBox.Show("Please enter customer name", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCustomerName.Focus()
            Return False
        End If
        
        Return True
    End Function
    
    Private Sub ClearCustomerForm()
        txtCustomerName.Clear()
        txtContactNumber.Clear()
        txtEmail.Clear()
        txtAddress.Clear()
        editingCustomerID = -1
    End Sub
    
    Private Sub btnDeleteCustomer_Click(sender As Object, e As EventArgs) Handles btnDeleteCustomer.Click
        If dgvCustomers.SelectedRows.Count > 0 Then
            If MessageBox.Show("Are you sure you want to delete this customer?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    Dim customerID As Integer = CInt(dgvCustomers.SelectedRows(0).Cells("CustomerID").Value)
                    
                    Using conn As New OleDbConnection(dbConnection)
                        conn.Open()
                        
                        Dim cmd As New OleDbCommand("DELETE FROM tblCustomers WHERE CustomerID = @id", conn)
                        cmd.Parameters.AddWithValue("@id", customerID)
                        
                        cmd.ExecuteNonQuery()
                        
                        MessageBox.Show("Customer deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ClearCustomerForm()
                        LoadCustomers()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error deleting customer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a customer to delete", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearCustomerForm()
    End Sub
End Class