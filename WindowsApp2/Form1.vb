Imports System.Data.OleDb

Public Class FORM1

    Dim conn As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim dt As DataTable
    Dim da As OleDbDataAdapter

    Private bitmap As Bitmap

    Private Sub viewer()
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()

        conn.Open()
        cmd = conn.CreateCommand()
        cmd.CommandType = CommandType.Text
        da = New OleDbDataAdapter("select * FROM PRODUCTS", conn)
        dt = New DataTable()
        da.Fill(dt)
        DataGridView1.DataSource = dt
        DataGridView2.DataSource = dt
        conn.Close()


        DataGridView1.Columns(0).Width = 130
        DataGridView1.Columns(1).Width = 130
        DataGridView1.Columns(2).Width = 130
        DataGridView1.Columns(3).Width = 130
        DataGridView1.Columns(4).Width = 130
        DataGridView1.Columns(5).Width = 130
    End Sub

    Private Sub btnPassword_Click(sender As Object, e As EventArgs)
        GroupBox1.Visible = False
    End Sub

    Private Sub RESET_PW_BTN_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub EXIT_MGMT_BTN_Click(sender As Object, e As EventArgs)
        Dim result = MessageBox.Show("Are you sure you would like to exit?", "Closing Exams", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GroupBox1.Visible = False
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
    End Sub

    Private Sub ADMINBUTTON_Click(sender As Object, e As EventArgs) Handles ADMINBUTTON.Click
        GroupBox1.Visible = True
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
    End Sub

    Private deletedRow As DataRow

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles DELETE_BTN.Click

        If DataGridView1.SelectedRows.Count > 0 Then
            Try
                ' Store the selected row data for undo
                deletedRow = CType(DataGridView1.SelectedRows(0).DataBoundItem, DataRowView).Row

                ' Delete the selected row from the database
                conn.Open()
                cmd = conn.CreateCommand()
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "DELETE FROM PRODUCTS WHERE ALL_ID = @allId"
                cmd.Parameters.AddWithValue("@allId", deletedRow("ALL_ID").ToString())
                cmd.ExecuteNonQuery()
                conn.Close()

                ' Refresh the DataGridView
                viewer()
                MessageBox.Show("Record Deleted Successfully", "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        Else
            MessageBox.Show("No record selected to delete", "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
    End Sub

    Private Sub FORM1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tinee\OneDrive\Documents\product_catalog.accdb"
        viewer()
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
    End Sub

    Private Sub ADD_BTN_Click(sender As Object, e As EventArgs) Handles ADD_BTN.Click

        Try

            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into PRODUCTS(ALL_ID,ALL_NAME,ALL_CATEGORY,ALL_BRAND,ALL_SIZES,ALL_BRANCHES)values('" + txtID.Text + "', '" + txtName.Text + "', '" + txtCategory.Text + "', '" + txtBrand.Text + "', '" + txtSize.Text + "','" + txtBranch.Text + "')"
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Record Saved Successfully in MS Access", "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Information)

            viewer()

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub

    Private Sub VIEW_BTN_Click(sender As Object, e As EventArgs) Handles VIEW_BTN.Click

        viewer()

    End Sub

    Private Sub UPDATE_BTN_Click(sender As Object, e As EventArgs) Handles UPDATE_BTN.Click

        Try

            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "UPDATE PRODUCTS SET ALL_ID = '" + txtID.Text + "', ALL_NAME = '" + txtName.Text + "' , ALL_CATEGORY = '" + txtCategory.Text + "', ALL_BRAND = '" + txtBrand.Text + "', ALL_SIZES = '" + txtSize.Text + "' where ALL_BRANCHES = '" + txtBranch.Text + "'"
            cmd.ExecuteNonQuery()
            conn.Close()
            MessageBox.Show("Record Saved Successfully in MS Access", "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Information)

            viewer()

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try

            txtID.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString()
            txtName.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString()
            txtCategory.Text = DataGridView1.SelectedRows(0).Cells(2).Value.ToString()
            txtBrand.Text = DataGridView1.SelectedRows(0).Cells(3).Value.ToString()
            txtSize.Text = DataGridView1.SelectedRows(0).Cells(4).Value.ToString()
            txtBranch.Text = DataGridView1.SelectedRows(0).Cells(5).Value.ToString()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UNDO_DELETE_BTN_Click(sender As Object, e As EventArgs) Handles UNDO_DELETE_BTN.Click
        If deletedRow IsNot Nothing Then
            Try
                ' Reinsert the deleted row into the database
                conn.Open()
                cmd = conn.CreateCommand()
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "INSERT INTO PRODUCTS (ALL_ID, ALL_NAME, ALL_CATEGORY, ALL_BRAND, ALL_SIZES, ALL_BRANCHES) VALUES (@allId, @allName, @allCategory, @allBrand, @allSizes, @allBranches)"
                cmd.Parameters.AddWithValue("@allId", deletedRow("ALL_ID").ToString())
                cmd.Parameters.AddWithValue("@allName", deletedRow("ALL_NAME").ToString())
                cmd.Parameters.AddWithValue("@allCategory", deletedRow("ALL_CATEGORY").ToString())
                cmd.Parameters.AddWithValue("@allBrand", deletedRow("ALL_BRAND").ToString())
                cmd.Parameters.AddWithValue("@allSizes", deletedRow("ALL_SIZES").ToString())
                cmd.Parameters.AddWithValue("@allBranches", deletedRow("ALL_BRANCHES").ToString())
                cmd.ExecuteNonQuery()
                conn.Close()

                ' Clear the stored deleted row
                deletedRow = Nothing

                ' Refresh the DataGridView
                viewer()
                MessageBox.Show("Undo successful, record restored", "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        Else
            MessageBox.Show("No record to undo", "VB Save Database", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Private Sub RESET_BTN_Click(sender As Object, e As EventArgs) Handles RESET_BTN.Click


        txtID.Text = " "
        txtName.Text = " "
        txtCategory.Text = " "
        txtBrand.Text = " "
        txtSize.Text = " "
        txtBranch.Text = " "

    End Sub

    Private Sub SEARCH_BTN_Click(sender As Object, e As EventArgs) Handles SEARCH_BTN.Click

        Dim checker As Integer


        Try

            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM PRODUCTS WHERE ALL_ID = '" + txtID.Text + "' OR ALL_NAME = '" + txtName.Text.Trim() + "' OR ALL_CATEGORY = '" + txtCategory.Text.Trim() + "' OR ALL_BRAND = '" + txtBrand.Text.Trim() + "' OR ALL_SIZES = '" + txtSize.Text.Trim() + "' OR ALL_BRANCHES = '" + txtBranch.Text + "'"
            cmd.ExecuteNonQuery()
            dt = New DataTable()
            da = New OleDbDataAdapter(cmd)
            da.Fill(dt)

            checker = Convert.ToInt32(dt.Rows.Count.ToString())
            DataGridView1.DataSource = dt
            conn.Close()
            If (checker = 0) Then
                txtSearchh.Text = "Search"
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub SEARCH_BUTTON_Click(sender As Object, e As EventArgs) Handles SEARCH_BUTTON_USER.Click

        Dim checker As Integer


        Try

            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM PRODUCTS WHERE ALL_NAME = '" + txtName2.Text.Trim() + "' OR ALL_CATEGORY = '" + TxtCategory2.Text.Trim() + "' OR ALL_BRAND = '" + txtBrand2.Text.Trim() + "' OR ALL_SIZES = '" + txtSizes2.Text.Trim() + "' OR ALL_BRANCHES = '" + txtBranch2.Text + "'"
            cmd.ExecuteNonQuery()
            dt = New DataTable()
            da = New OleDbDataAdapter(cmd)
            da.Fill(dt)

            checker = Convert.ToInt32(dt.Rows.Count.ToString())
            DataGridView2.DataSource = dt
            conn.Close()
            If (checker = 0) Then
                txtSearchh.Text = "Search"
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TxtCategory2.TextChanged

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub txtName2_TextChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub txtName2_TextChanged_1(sender As Object, e As EventArgs) Handles txtName2.TextChanged
    End Sub

    Private Sub VIEW_BTN_USER_Click(sender As Object, e As EventArgs) Handles VIEW_BTN_USER.Click
        viewer()
    End Sub
End Class
