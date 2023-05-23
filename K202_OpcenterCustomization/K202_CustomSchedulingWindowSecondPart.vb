Public Class K202_CustomSchedulingWindowSecondPart
    Public dataTable As DataTable = New DataTable()
    Public okBtnClicked As Boolean = False
    Private Sub CustomSchedulingWindowSecondPart_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = dataTable
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim columnName As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name
    End Sub

    Private Sub Ok_Btn_Click(sender As Object, e As EventArgs) Handles Ok_Btn.Click
        okBtnClicked = True
        Close()
    End Sub
End Class