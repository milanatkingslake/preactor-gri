
Imports System.Windows.Forms
Public Class K202_JobSplitDetails
    Public Property strJobOrderNo As String
    Public Property decJobOrderQty As Double
    Public Property isOkClick As Boolean
    Public Property decOriginalJobOrderQty As Double
    Private Sub K203_JobSplitDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        JobOrderNoTxt.Text = strJobOrderNo
        JobQtyTxt.Text = decJobOrderQty.ToString
        decOriginalJobOrderQty = decJobOrderQty
    End Sub

    Private Sub K203_JobSplitDetails_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If isOkClick Then
            sender = JobQtyTxt.Text
        Else
            sender = ""
        End If
    End Sub

    Private Sub Ok_Click(sender As Object, e As EventArgs) Handles OkBtn.Click
        If decOriginalJobOrderQty > CDec(JobQtyTxt.Text) Then
            isOkClick = True
            Close()
        Else
            MsgBox("Split Quantity should be less than the parent job quantity",, "Information")
        End If
        'isOkClick = True
        'Close()
    End Sub

    Private Sub QuantityTxt_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub JobQtyTxt_KeyPress(sender As Object, e As Windows.Forms.KeyPressEventArgs) Handles JobQtyTxt.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) AndAlso (e.KeyChar <> "."c) Then
            e.Handled = True
        End If

        If (e.KeyChar = "."c) AndAlso ((TryCast(sender, TextBox)).Text.IndexOf("."c) > -1) Then
            e.Handled = True
        End If
    End Sub

    Private Sub JobOrderNoTxt_TextChanged(sender As Object, e As EventArgs) Handles JobOrderNoTxt.TextChanged

    End Sub
End Class