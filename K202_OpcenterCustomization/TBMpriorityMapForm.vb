Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class TBMpriorityMapForm
    'Public Property tblFormerDetailsMain As DataTable
    'Public Property tblSize As DataTable
    'Public Property tblOrder As DataTable
    'Public Property tbltblOrderRate_gl As DataTable
    Public Property connetionString As String
    Private Sub TBMpriorityMapForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBoxRule.SelectedIndex = 1
        Load_DataGridViewAllResourceGroup()
        Load_DataGridViewMapResourceGroup()
    End Sub
    Sub Load_DataGridViewAllResourceGroup()
        Dim con As SqlConnection = New SqlConnection(connetionString)
        ''Dim cmd As SqlCommand = New SqlCommand("select 0 'Checked',name 'Group' from userdata.ResourceGroups where name not in (select GroupCode 'Group' from dbo.K202_TBMPriorityMapDetails where RuleCode= " + "'" + ComboBoxRule.SelectedItem.ToString + "')", con)
        Dim cmd As SqlCommand = New SqlCommand("select 0 'Checked',name 'Group',ResourceGroupsId from userdata.ResourceGroups where name not in (select GroupCode 'Group' from dbo.K202_TBMPriorityMapDetails )", con)

        Dim Adpt As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim dt As DataTable = New DataTable()
        Adpt.Fill(dt)

        DataGridViewAllResourceGroup.AutoGenerateColumns = False
        DataGridViewAllResourceGroup.Columns.Clear()

        Dim column As DataGridViewColumn = New DataGridViewCheckBoxColumn()
        column.DataPropertyName = "Checked"
        column.Name = "Checked"
        DataGridViewAllResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "Group"
        column.Name = "Group"
        column.Width = 200
        DataGridViewAllResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "ResourceGroupsId"
        column.Name = "ResourceGroupsId"
        column.Width = 50
        column.Visible = False
        DataGridViewAllResourceGroup.Columns.Add(column)

        Me.Controls.Add(DataGridViewAllResourceGroup)
        Me.AutoSize = True
        Me.Text = ""
        DataGridViewAllResourceGroup.AutoSize = True
        DataGridViewAllResourceGroup.DataSource = dt
        DataGridViewAllResourceGroup.AllowUserToAddRows = False

    End Sub
    Sub Load_DataGridViewMapResourceGroup()
        Dim con As SqlConnection = New SqlConnection(connetionString)
        Dim cmd As SqlCommand = New SqlCommand(" select 0 'Checked' , GroupCode 'Group',ResourceGroupsId from dbo.K202_TBMPriorityMapDetails where RuleCode= " + "'" + ComboBoxRule.SelectedItem.ToString + "'", con)
        Dim Adpt As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim dt As DataTable = New DataTable()
        Adpt.Fill(dt)
        DataGridViewMapResourceGroup.Columns.Clear()
        DataGridViewMapResourceGroup.AutoGenerateColumns = False

        Dim column As DataGridViewColumn = New DataGridViewCheckBoxColumn()
        column.DataPropertyName = "Checked"
        column.Name = "Checked"
        DataGridViewMapResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "Group"
        column.Name = "Group"
        column.Width = 200
        DataGridViewMapResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "ResourceGroupsId"
        column.Name = "ResourceGroupsId"
        column.Width = 50
        column.Visible = False
        DataGridViewMapResourceGroup.Columns.Add(column)

        Me.Controls.Add(DataGridViewMapResourceGroup)
        Me.AutoSize = True
        Me.Text = ""
        ''DataGridViewMapResourceGroup.AutoSize = True
        DataGridViewMapResourceGroup.DataSource = dt
        DataGridViewMapResourceGroup.AllowUserToAddRows = False

    End Sub
    Private Sub btnMap_Click(sender As Object, e As EventArgs) Handles btnMap.Click
        'Dim message As String = String.Empty
        'For Each row As DataGridViewRow In DataGridViewAllResourceGroup.Rows
        '    Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
        '    If isSelected Then
        '        message &= Environment.NewLine
        '        message &= row.Cells("Group").Value.ToString()
        '    End If
        'Next
        'MessageBox.Show("Selected Values" & message)
        'DataGridViewMapResourceGroup.Rows.Clear()
        'DataGridViewMapResourceGroup.Refresh()

        ''Map Grid view

        Dim dtMap As DataTable = New DataTable()
        Dim checked As DataColumn = New DataColumn("Checked", Type.[GetType]("System.String"))
        Dim Group As DataColumn = New DataColumn("Group", Type.[GetType]("System.String"))
        Dim ResourceGroupsId As DataColumn = New DataColumn("ResourceGroupsId", Type.[GetType]("System.String"))

        dtMap.Columns.Add(checked)
        dtMap.Columns.Add(Group)
        dtMap.Columns.Add(ResourceGroupsId)

        For Each row As DataGridViewRow In DataGridViewMapResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
                If Not isSelected Then
                    Dim dr_int As DataRow = dtMap.NewRow()
                    dr_int("Checked") = "False"
                    dr_int("Group") = row.Cells("Group").Value.ToString()
                    dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                    dtMap.Rows.Add(dr_int)
                End If

            End If
        Next

        For Each row As DataGridViewRow In DataGridViewAllResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
                If isSelected Then
                    Dim dr_int As DataRow = dtMap.NewRow()
                    dr_int("Checked") = "False"
                    dr_int("Group") = row.Cells("Group").Value.ToString()
                    dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                    dtMap.Rows.Add(dr_int)
                End If
            End If
        Next
        DataGridViewMapResourceGroup.Columns.Clear()
        DataGridViewMapResourceGroup.AutoGenerateColumns = False
        If Not DataGridViewMapResourceGroup.Columns.Count > 0 Then
            Dim column As DataGridViewColumn = New DataGridViewCheckBoxColumn()
            column.DataPropertyName = "Checked"
            column.Name = "Checked"
            DataGridViewMapResourceGroup.Columns.Add(column)

            column = New DataGridViewTextBoxColumn()
            column.DataPropertyName = "Group"
            column.Name = "Group"
            column.Width = 200
            DataGridViewMapResourceGroup.Columns.Add(column)

            column = New DataGridViewTextBoxColumn()
            column.DataPropertyName = "ResourceGroupsId"
            column.Name = "ResourceGroupsId"
            column.Width = 50
            column.Visible = False
            DataGridViewMapResourceGroup.Columns.Add(column)

            Me.Controls.Add(DataGridViewMapResourceGroup)
        End If
        Me.AutoSize = True
        Me.Text = ""
        DataGridViewMapResourceGroup.AutoSize = True
        DataGridViewMapResourceGroup.DataSource = dtMap

        ''All Grid view
        Dim dtAll As DataTable = New DataTable()
        Dim checkedAll As DataColumn = New DataColumn("Checked", Type.[GetType]("System.String"))
        Dim GroupAll As DataColumn = New DataColumn("Group", Type.[GetType]("System.String"))
        Dim ResourceGroupsIdAll As DataColumn = New DataColumn("ResourceGroupsId", Type.[GetType]("System.String"))


        dtAll.Columns.Add(checkedAll)
        dtAll.Columns.Add(GroupAll)
        dtAll.Columns.Add(ResourceGroupsIdAll)


        For Each row As DataGridViewRow In DataGridViewAllResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
                If Not isSelected Then
                    Dim dr_int As DataRow = dtAll.NewRow()
                    dr_int("Checked") = "False"
                    dr_int("Group") = row.Cells("Group").Value.ToString()
                    dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                    dtAll.Rows.Add(dr_int)
                End If
            End If
        Next
        DataGridViewAllResourceGroup.Columns.Clear()
        DataGridViewAllResourceGroup.AutoGenerateColumns = False
        If Not DataGridViewAllResourceGroup.Columns.Count > 0 Then
            Dim column As DataGridViewColumn = New DataGridViewCheckBoxColumn()
            column.DataPropertyName = "Checked"
            column.Name = "Checked"
            DataGridViewAllResourceGroup.Columns.Add(column)

            column = New DataGridViewTextBoxColumn()
            column.DataPropertyName = "Group"
            column.Name = "Group"
            column.Width = 200
            DataGridViewAllResourceGroup.Columns.Add(column)

            column = New DataGridViewTextBoxColumn()
            column.DataPropertyName = "ResourceGroupsId"
            column.Name = "ResourceGroupsId"
            column.Width = 50
            column.Visible = False
            DataGridViewAllResourceGroup.Columns.Add(column)

            Me.Controls.Add(DataGridViewAllResourceGroup)
        End If
        Me.AutoSize = True
        Me.Text = ""
        DataGridViewAllResourceGroup.AutoSize = True
        DataGridViewAllResourceGroup.DataSource = dtAll
    End Sub
    Private Sub btnUnMap_Click(sender As Object, e As EventArgs) Handles btnUnMap.Click
        Dim dtAll As DataTable = New DataTable()
        Dim dtMap As DataTable = New DataTable()


        Dim checkedAll As DataColumn = New DataColumn("Checked", Type.[GetType]("System.String"))
        Dim GroupAll As DataColumn = New DataColumn("Group", Type.[GetType]("System.String"))
        Dim ResourceGroupsIdAll As DataColumn = New DataColumn("ResourceGroupsId", Type.[GetType]("System.String"))

        Dim checkedMap As DataColumn = New DataColumn("Checked", Type.[GetType]("System.String"))
        Dim GroupMap As DataColumn = New DataColumn("Group", Type.[GetType]("System.String"))
        Dim ResourceGroupsIdMap As DataColumn = New DataColumn("ResourceGroupsId", Type.[GetType]("System.String"))

        dtAll.Columns.Add(checkedAll)
        dtAll.Columns.Add(GroupAll)
        dtAll.Columns.Add(ResourceGroupsIdAll)

        dtMap.Columns.Add(checkedMap)
        dtMap.Columns.Add(GroupMap)
        dtMap.Columns.Add(ResourceGroupsIdMap)



        For Each row As DataGridViewRow In DataGridViewAllResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim dr_int As DataRow = dtAll.NewRow()
                dr_int("Checked") = "False"
                dr_int("Group") = row.Cells("Group").Value.ToString()
                dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                dtAll.Rows.Add(dr_int)
            End If
        Next

        For Each row As DataGridViewRow In DataGridViewMapResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
                If isSelected And Not (row.Cells("Group").Value.ToString() = "") Then
                    Dim dr_int As DataRow = dtAll.NewRow()
                    dr_int("Checked") = "False"
                    dr_int("Group") = row.Cells("Group").Value.ToString()
                    dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                    dtAll.Rows.Add(dr_int)
                End If
            End If
        Next

        DataGridViewAllResourceGroup.AutoGenerateColumns = False
        DataGridViewAllResourceGroup.Columns.Clear()
        DataGridViewAllResourceGroup.AllowUserToAddRows = False
        Dim column As DataGridViewColumn = New DataGridViewCheckBoxColumn()
        column.DataPropertyName = "Checked"
        column.Name = "Checked"
        DataGridViewAllResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "Group"
        column.Name = "Group"
        column.Width = 200
        DataGridViewAllResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "ResourceGroupsId"
        column.Name = "ResourceGroupsId"
        column.Width = 50
        column.Visible = False
        DataGridViewAllResourceGroup.Columns.Add(column)

        Me.Controls.Add(DataGridViewAllResourceGroup)
        Me.AutoSize = True
        Me.Text = ""
        DataGridViewAllResourceGroup.AutoSize = True

        DataGridViewAllResourceGroup.DataSource = dtAll
        '''dtMap data adding
        For Each row As DataGridViewRow In DataGridViewMapResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Checked").Value)
                If Not isSelected Then
                    Dim dr_int As DataRow = dtMap.NewRow()
                    dr_int("Checked") = "False"
                    dr_int("Group") = row.Cells("Group").Value.ToString()
                    dr_int("ResourceGroupsId") = row.Cells("ResourceGroupsId").Value.ToString()

                    dtMap.Rows.Add(dr_int)
                End If
            End If
        Next

        DataGridViewMapResourceGroup.AutoGenerateColumns = False
        DataGridViewMapResourceGroup.Columns.Clear()
        column = New DataGridViewCheckBoxColumn()
        column.DataPropertyName = "Checked"
        column.Name = "Checked"
        DataGridViewMapResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "Group"
        column.Name = "Group"
        column.Width = 200
        DataGridViewMapResourceGroup.Columns.Add(column)

        column = New DataGridViewTextBoxColumn()
        column.DataPropertyName = "ResourceGroupsId"
        column.Name = "ResourceGroupsId"
        column.Width = 50
        column.Visible = False
        DataGridViewMapResourceGroup.Columns.Add(column)

        Me.Controls.Add(DataGridViewMapResourceGroup)
        Me.AutoSize = True
        Me.Text = ""
        DataGridViewMapResourceGroup.AutoSize = True
        DataGridViewMapResourceGroup.DataSource = dtMap
        DataGridViewAllResourceGroup.AllowUserToAddRows = False
        DataGridViewMapResourceGroup.AllowUserToAddRows = False

    End Sub
    Private Sub Save_Click(sender As Object, e As EventArgs) Handles Save.Click
        Dim strMapGroups As String = ""
        For Each row As DataGridViewRow In DataGridViewMapResourceGroup.Rows
            If Not IsNothing(row.Cells("Checked").Value) Then
                strMapGroups = strMapGroups + row.Cells("ResourceGroupsId").Value.ToString() + ","
            End If
        Next
        If strMapGroups = "" Then
            MsgBox("No Map Record Selected unble to save...")
        Else
            strMapGroups = strMapGroups.Remove(strMapGroups.Length - 1)
            UpdateMapGroups(ComboBoxRule.SelectedItem.ToString(), strMapGroups)
        End If

    End Sub
    Public Function UpdateMapGroups(ByRef strRule As String, ByRef strMapGroupIds As String) As Boolean

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand
            Dim status As Boolean
            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_TBMRuleReosurceGroupMap_Sp"
            Dim param As SqlParameter

            param = New SqlParameter("@StrRule", strRule)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@strMapGroupIds", strMapGroupIds)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            command.CommandTimeout = 340
            command.ExecuteNonQuery()

            Return True
            connection.Close()
        Catch ex As Exception
            Return False
            MsgBox("Extra formers not define",, "error")
        Finally
            MsgBox("Record updated...")
        End Try

    End Function

    Private Sub ComboBoxRule_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxRule.SelectedIndexChanged
        Load_DataGridViewAllResourceGroup()
        Load_DataGridViewMapResourceGroup()
    End Sub
End Class