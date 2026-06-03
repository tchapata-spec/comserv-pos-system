Imports System.Data.OleDb

Public Class PromotionManagement
    Private dbConnection As String = GlobalVariables.DBConnectionString
    Private editingPromotionID As Integer = -1
    
    Private Sub PromotionManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Promotion Management"
        LoadPromotions()
    End Sub
    
    Private Sub LoadPromotions()
        Try
            Using conn As New OleDbConnection(dbConnection)
                conn.Open()
                
                Dim cmd As New OleDbCommand("SELECT PromotionID, PromotionName, DiscountPercentage, StartDate, EndDate, IsActive FROM tblPromotions ORDER BY StartDate DESC", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)
                
                dgvPromotions.DataSource = table
                dgvPromotions.AutoResizeColumns()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading promotions: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnAddPromotion_Click(sender As Object, e As EventArgs) Handles btnAddPromotion.Click
        If ValidatePromotionInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("INSERT INTO tblPromotions (PromotionName, DiscountPercentage, StartDate, EndDate, ApplicableProducts, MinimumPurchase, IsActive) VALUES (@name, @discount, @start, @end, @products, @minPurchase, @active)", conn)
                    cmd.Parameters.AddWithValue("@name", txtPromotionName.Text)
                    cmd.Parameters.AddWithValue("@discount", CDec(txtDiscountPercentage.Text))
                    cmd.Parameters.AddWithValue("@start", dtpStartDate.Value)
                    cmd.Parameters.AddWithValue("@end", dtpEndDate.Value)
                    cmd.Parameters.AddWithValue("@products", txtApplicableProducts.Text)
                    cmd.Parameters.AddWithValue("@minPurchase", If(String.IsNullOrWhiteSpace(txtMinimumPurchase.Text), 0, CDec(txtMinimumPurchase.Text)))
                    cmd.Parameters.AddWithValue("@active", chkActive.Checked)
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Promotion added successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearPromotionForm()
                    LoadPromotions()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error adding promotion: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub btnUpdatePromotion_Click(sender As Object, e As EventArgs) Handles btnUpdatePromotion.Click
        If editingPromotionID = -1 Then
            MessageBox.Show("Please select a promotion to update", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        
        If ValidatePromotionInput() Then
            Try
                Using conn As New OleDbConnection(dbConnection)
                    conn.Open()
                    
                    Dim cmd As New OleDbCommand("UPDATE tblPromotions SET PromotionName = @name, DiscountPercentage = @discount, StartDate = @start, EndDate = @end, ApplicableProducts = @products, MinimumPurchase = @minPurchase, IsActive = @active WHERE PromotionID = @id", conn)
                    cmd.Parameters.AddWithValue("@name", txtPromotionName.Text)
                    cmd.Parameters.AddWithValue("@discount", CDec(txtDiscountPercentage.Text))
                    cmd.Parameters.AddWithValue("@start", dtpStartDate.Value)
                    cmd.Parameters.AddWithValue("@end", dtpEndDate.Value)
                    cmd.Parameters.AddWithValue("@products", txtApplicableProducts.Text)
                    cmd.Parameters.AddWithValue("@minPurchase", If(String.IsNullOrWhiteSpace(txtMinimumPurchase.Text), 0, CDec(txtMinimumPurchase.Text)))
                    cmd.Parameters.AddWithValue("@active", chkActive.Checked)
                    cmd.Parameters.AddWithValue("@id", editingPromotionID)
                    
                    cmd.ExecuteNonQuery()
                    
                    MessageBox.Show("Promotion updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ClearPromotionForm()
                    editingPromotionID = -1
                    LoadPromotions()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error updating promotion: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    
    Private Sub dgvPromotions_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPromotions.CellDoubleClick
        If e.RowIndex >= 0 Then
            editingPromotionID = CInt(dgvPromotions.Rows(e.RowIndex).Cells("PromotionID").Value)
            txtPromotionName.Text = dgvPromotions.Rows(e.RowIndex).Cells("PromotionName").Value.ToString()
            txtDiscountPercentage.Text = dgvPromotions.Rows(e.RowIndex).Cells("DiscountPercentage").Value.ToString()
            dtpStartDate.Value = CDate(dgvPromotions.Rows(e.RowIndex).Cells("StartDate").Value)
            dtpEndDate.Value = CDate(dgvPromotions.Rows(e.RowIndex).Cells("EndDate").Value)
            chkActive.Checked = CBool(dgvPromotions.Rows(e.RowIndex).Cells("IsActive").Value)
        End If
    End Sub
    
    Private Function ValidatePromotionInput() As Boolean
        If String.IsNullOrWhiteSpace(txtPromotionName.Text) Then
            MessageBox.Show("Please enter promotion name", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPromotionName.Focus()
            Return False
        End If
        
        If Not Decimal.TryParse(txtDiscountPercentage.Text, Nothing) Then
            MessageBox.Show("Please enter a valid discount percentage", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDiscountPercentage.Focus()
            Return False
        End If
        
        Dim discount As Decimal = CDec(txtDiscountPercentage.Text)
        If discount < 0 Or discount > 100 Then
            MessageBox.Show("Discount percentage must be between 0 and 100", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDiscountPercentage.Focus()
            Return False
        End If
        
        If dtpStartDate.Value >= dtpEndDate.Value Then
            MessageBox.Show("Start date must be before end date", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        
        Return True
    End Function
    
    Private Sub ClearPromotionForm()
        txtPromotionName.Clear()
        txtDiscountPercentage.Clear()
        txtApplicableProducts.Clear()
        txtMinimumPurchase.Clear()
        dtpStartDate.Value = DateTime.Now
        dtpEndDate.Value = DateTime.Now.AddDays(30)
        chkActive.Checked = True
        editingPromotionID = -1
    End Sub
    
    Private Sub btnDeletePromotion_Click(sender As Object, e As EventArgs) Handles btnDeletePromotion.Click
        If dgvPromotions.SelectedRows.Count > 0 Then
            If MessageBox.Show("Are you sure you want to delete this promotion?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    Dim promotionID As Integer = CInt(dgvPromotions.SelectedRows(0).Cells("PromotionID").Value)
                    
                    Using conn As New OleDbConnection(dbConnection)
                        conn.Open()
                        
                        Dim cmd As New OleDbCommand("DELETE FROM tblPromotions WHERE PromotionID = @id", conn)
                        cmd.Parameters.AddWithValue("@id", promotionID)
                        
                        cmd.ExecuteNonQuery()
                        
                        MessageBox.Show("Promotion deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ClearPromotionForm()
                        LoadPromotions()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error deleting promotion: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a promotion to delete", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearPromotionForm()
    End Sub
End Class