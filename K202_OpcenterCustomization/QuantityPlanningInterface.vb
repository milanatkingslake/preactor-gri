Imports System.Drawing

Public Class QuantityPlanningInterface
    Public dataTable As DataTable = New DataTable()
    Public Property saveBtnClicked As Boolean = False
    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim columnName As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name
        If columnName = "Order Qty CLM" Then
            If Not (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then

                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value) - CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value)
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").ReadOnly = True
                Else
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            ElseIf (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value)
                Else
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            ElseIf (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") Then
                If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value)
                Else
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            Else
                DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
            End If
        End If

        If columnName = "Access Qty" Then
            If Not DataGridView1.CurrentRow.Cells("Order Qty CLM").Value.ToString = "" Then
                If Not (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                    If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then

                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value) - CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value)
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").ReadOnly = True
                    Else
                        DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                ElseIf (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") Then
                    If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then
                        DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value)
                    Else
                        DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                ElseIf (DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CMP").Value.ToString = "0") And Not (DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "" Or DataGridView1.CurrentRow.Cells("No Of CLM").Value.ToString = "0") Then
                    If CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) >= CInt(DataGridView1.CurrentRow.Cells("Order Qty CLM").Value) Then
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                        DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = CInt(DataGridView1.CurrentRow.Cells("Planning Qty").Value) + CInt(DataGridView1.CurrentRow.Cells("Access Qty").Value)
                    Else
                        DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                        DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                    End If
                Else
                    DataGridView1.CurrentRow.Cells("Order Qty CLM").Value = 0
                    DataGridView1.CurrentRow.Cells("Order Qty CMP").Value = 0
                End If
            End If
        End If


    End Sub

    Private Sub QuantityPlanningInterface_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = dataTable
        DataGridView1.Columns("Urgency").DefaultCellStyle.Alignment = Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridView1.Columns("Urgency").DefaultCellStyle.ForeColor = Color.Red
        DataGridView1.Columns("Urgency").DefaultCellStyle.Font = New Font("Tahoma", 9.25, FontStyle.Bold)
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim k As New DataView(dataTable)
        k.RowFilter = String.Format("Sku like '%{0}%'", TextBox1.Text)
        DataGridView1.DataSource = k
    End Sub

End Class