Imports System.Drawing

Public Class QuantityPlanningInterface
    Public tblOrders As DataTable = New DataTable()

    Public Property saveBtnClicked As Boolean = False
    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles GridViewMasterProduction.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles GridViewMasterProduction.CellEndEdit
        Dim columnName As String = GridViewMasterProduction.Columns(GridViewMasterProduction.CurrentCell.ColumnIndex).Name
        If columnName = "Order Qty CLM" Then
            If Not (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then

                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value) - CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value)
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").ReadOnly = True
                Else
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            ElseIf (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value)
                Else
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            ElseIf (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") Then
                If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value)
                Else
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            Else
                GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
            End If
        End If

        If columnName = "Access Qty" Then
            If Not GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value.ToString = "" Then
                If Not (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                    If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then

                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value) - CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value)
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").ReadOnly = True
                    Else
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                ElseIf (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                    If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value)
                    Else
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                ElseIf (GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CMP").Value.ToString = "0") And Not (GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or GridViewMasterProduction.CurrentRow.Cells("No Of CLM").Value.ToString = "0") Then
                    If CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) >= CInt(GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value) Then
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = CInt(GridViewMasterProduction.CurrentRow.Cells("Planning Qty").Value) + CInt(GridViewMasterProduction.CurrentRow.Cells("Access Qty").Value)
                    Else
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                        GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                Else
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CLM").Value = 0
                    GridViewMasterProduction.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            End If
        End If


    End Sub

    Private Sub QuantityPlanningInterface_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GridViewMasterProduction.DataSource = tblOrders

        'dt.Columns("SKU").ReadOnly = True
        'dt.Columns("Description").ReadOnly = True
        'dt.Columns("Planning Qty").ReadOnly = True
        'dt.Columns("No Of CLM").ReadOnly = True
        'dt.Columns("No Of CMP").ReadOnly = True
        'dt.Columns("On Hand").ReadOnly = True
        'dt.Columns("On Plan").ReadOnly = True
        'dt.Columns("Running Qty").ReadOnly = True


        'DataGridView1.Columns("SKU").ReadOnly = True
        'DataGridView1.Columns("Description").ReadOnly = True
        'DataGridView1.Columns("ExcessQty").ReadOnly = True
        'DataGridView1.Columns("PlanningQty").ReadOnly = True
        'DataGridView1.Columns("NoOfCLM").ReadOnly = True
        'DataGridView1.Columns("AvailableCLM").ReadOnly = True
        'DataGridView1.Columns("NoOfCMP").ReadOnly = True
        'DataGridView1.Columns("AvailableCMP").ReadOnly = True
        'DataGridView1.Columns("OrderQtyCLM").ReadOnly = True
        'DataGridView1.Columns("OrderQtyCMP").ReadOnly = True
        'DataGridView1.Columns("BalanceQty").ReadOnly = True
        'DataGridView1.Columns("Urgency").ReadOnly = True

        GridViewMasterProduction.ReadOnly = True

        GridViewMasterProduction.Columns("Urgency").DefaultCellStyle.Alignment = Windows.Forms.DataGridViewContentAlignment.TopCenter
        GridViewMasterProduction.Columns("Urgency").DefaultCellStyle.ForeColor = Color.Red
        GridViewMasterProduction.Columns("Urgency").DefaultCellStyle.Font = New Font("Tahoma", 9.25, FontStyle.Bold)
    End Sub

    'Private Sub SaveBtn_Click(sender As Object, e As EventArgs) Handles SaveBtn.Click
    '    saveBtnClicked = True
    '    'For c As Integer = 0 To DataGridView1.Rows.Count - 1
    '    '    If c >= 1 Then
    '    '        MsgBox(DataGridView1.Rows(c).Cells("Order Qty CMP").Value.ToString)
    '    '    End If

    '    'Next
    'End Sub
    Public Sub Close_Interface() Handles SaveBtn2.Click
        Close()
    End Sub

    Private Sub SaveBtn2_Click(sender As Object, e As EventArgs) Handles SaveBtn2.Click
        saveBtnClicked = True
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        Dim k As New DataView(tblOrders)
        k.RowFilter = String.Format("Sku like '%{0}%'", txtSKU.Text)
        GridViewMasterProduction.DataSource = k
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click
        Dim searchText As String = txtSalesReport.Text.Trim()
        Dim columnName As String

        Dim tblOrdersFilter As DataTable = New DataTable()

        tblOrdersFilter.Columns.Add(New DataColumn("Description", Type.[GetType]("System.String")))
        tblOrdersFilter.Columns.Add(New DataColumn("SKU", Type.[GetType]("System.String")))
        tblOrdersFilter.Columns.Add(New DataColumn("RejectedQty", Type.[GetType]("System.String")))
        tblOrdersFilter.Columns.Add(New DataColumn("OnHand", Type.[GetType]("System.String")))
        tblOrdersFilter.Columns.Add(New DataColumn("RunningQty", Type.[GetType]("System.String")))
        tblOrdersFilter.Columns.Add(New DataColumn("MonthlyDemand", Type.[GetType]("System.String")))


        If tblOrders.Columns.Count > 0 Then
            For i As Integer = 0 To tblOrders.Columns.Count - 1
                If Not String.IsNullOrEmpty(searchText) = True Then
                    columnName = tblOrders.Columns(i).ColumnName
                    If columnName = searchText Then
                        tblOrdersFilter.Columns.Add(New DataColumn(columnName, Type.[GetType]("System.String")))
                        Exit For
                    End If
                End If
            Next


            tblOrdersFilter.Columns.Add(New DataColumn("ExcessQty", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("PlanningQty", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("NoOfCLM", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("AvailableCLM", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("NoOfCMP", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("AvailableCMP", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("OrderQtyCLM", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("OrderQtyCMP", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("BalanceQty", Type.[GetType]("System.String")))
            tblOrdersFilter.Columns.Add(New DataColumn("Urgency", Type.[GetType]("System.String")))


            For Each row As DataRow In tblOrders.Rows
                If searchText <> "" Then
                    If (row(searchText).ToString()) <> "" Then
                        Dim newRow As DataRow = tblOrdersFilter.NewRow()
                        newRow("Description") = row("Description").ToString()
                        newRow("SKU") = row("SKU").ToString()
                        newRow("RejectedQty") = row("RejectedQty").ToString()
                        newRow("OnHand") = row("OnHand").ToString()
                        newRow("RunningQty") = row("RunningQty").ToString()
                        newRow("MonthlyDemand") = row("MonthlyDemand").ToString()
                        newRow(searchText) = row(searchText).ToString()
                        newRow("ExcessQty") = row("ExcessQty").ToString()
                        newRow("PlanningQty") = row("PlanningQty").ToString()
                        newRow("NoOfCLM") = row("NoOfCLM").ToString()
                        newRow("AvailableCLM") = row("AvailableCLM").ToString()
                        newRow("NoOfCMP") = row("NoOfCMP").ToString()
                        newRow("AvailableCMP") = row("AvailableCMP").ToString()
                        newRow("OrderQtyCLM") = row("OrderQtyCLM").ToString()
                        newRow("OrderQtyCMP") = row("OrderQtyCMP").ToString()
                        newRow("BalanceQty") = row("BalanceQty").ToString()
                        newRow("Urgency") = row("Urgency").ToString()



                        GridViewMasterProduction.Columns("ExcessQty").ReadOnly = True
                        GridViewMasterProduction.Columns("PlanningQty").ReadOnly = True
                        GridViewMasterProduction.Columns("NoOfCLM").ReadOnly = True
                        GridViewMasterProduction.Columns("AvailableCLM").ReadOnly = True
                        GridViewMasterProduction.Columns("NoOfCMP").ReadOnly = True
                        GridViewMasterProduction.Columns("AvailableCMP").ReadOnly = True
                        GridViewMasterProduction.Columns("OrderQtyCLM").ReadOnly = True
                        GridViewMasterProduction.Columns("OrderQtyCMP").ReadOnly = True
                        GridViewMasterProduction.Columns("BalanceQty").ReadOnly = True
                        GridViewMasterProduction.Columns("Urgency").ReadOnly = True



                        If Not String.IsNullOrEmpty(searchText) Then
                            newRow(searchText) = row(searchText).ToString()
                        End If
                        tblOrdersFilter.Rows.Add(newRow)
                    End If
                End If

            Next

        End If

        GridViewMasterProduction.DataSource = tblOrdersFilter
        GridViewMasterProduction.Refresh()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        GridViewMasterProduction.DataSource = tblOrders
        GridViewMasterProduction.Refresh()
    End Sub

    Private Sub txtSKU_TextChanged(sender As Object, e As EventArgs) Handles txtSKU.TextChanged
        Dim k As New DataView(tblOrders)
        k.RowFilter = String.Format("Sku like '%{0}%'", txtSKU.Text)
        GridViewMasterProduction.DataSource = k
    End Sub
End Class