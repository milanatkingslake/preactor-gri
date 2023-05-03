Imports System.Data.SqlClient

Public Class Form1
    Public d_table As DataTable = New DataTable()
    Dim connection As SqlConnection
    Dim selecteDSalesOrder As String = ""
    Dim selectedPartNo As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Sp_Call_Function(selecteDSalesOrder, selectedPartNo)
    End Sub

    Public Function form_sp_call(ByRef connetionString As String) As Integer
        connection = New SqlConnection(connetionString)
        Return 0
    End Function

    Private Sub SalesOrderComboBox_Control_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDataToComboBox("getSalesOrder_sp")
    End Sub

    Private Sub SKUComboBox_Control_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDataToComboBox("getPartNo_sp")
    End Sub

    Private Sub SalesOrderComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SalesOrderComboBox.SelectedIndexChanged
        selecteDSalesOrder = SalesOrderComboBox.SelectedItem.ToString()
        Sp_Call_Function(selecteDSalesOrder, selectedPartNo)
    End Sub

    Private Sub SKUComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SKUComboBox.SelectedIndexChanged
        selectedPartNo = SKUComboBox.SelectedItem.ToString()
        Sp_Call_Function(selecteDSalesOrder, selectedPartNo)
    End Sub

    Private Sub Sp_Call_Function(slsOrder As String, prtNo As String)

        Dim adapter As SqlDataAdapter
        Dim command As New SqlCommand
        connection.Open()
        command.Connection = connection
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "getOrdersByResourceName_Sp"

        Dim param As SqlParameter

        If (slsOrder = "") Then
            param = New SqlParameter("@salesOrder", DBNull.Value)
        Else
            param = New SqlParameter("@salesOrder", slsOrder)
        End If

        param.Direction = ParameterDirection.Input
        param.DbType = DbType.String
        command.Parameters.Add(param)

        If (prtNo = "") Then
            param = New SqlParameter("@partNo", DBNull.Value)
        Else
            param = New SqlParameter("@partNo", prtNo)
        End If
        param.Direction = ParameterDirection.Input
        param.DbType = DbType.String
        command.Parameters.Add(param)

        Dim da As New SqlDataAdapter
        da.SelectCommand = command
        Dim dt As New DataTable
        dt.Clear()
        da.Fill(dt)

        DataGridView.DataSource = dt
        connection.Close()
    End Sub

    Private Sub LoadDataToComboBox(sp_name As String)
        Dim adapter As SqlDataAdapter
        Dim command As New SqlCommand
        Dim ds2 As DataSet = New DataSet()
        connection.Open()
        command.Connection = connection
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = sp_name

        Dim da As New SqlDataAdapter
        da.SelectCommand = command
        Dim dt As New DataTable
        dt.Clear()
        da.Fill(dt)
        connection.Close()

        For Each row As DataRow In dt.Rows
            For Each column As DataColumn In dt.Columns
                If (sp_name = "getSalesOrder_sp") Then
                    SalesOrderComboBox.Items.Add(row(column))
                ElseIf (sp_name = "getPartNo_sp") Then
                    SKUComboBox.Items.Add(row(column))
                End If

            Next
        Next
    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub
End Class