Public Class K202_CustomSchedulingWindow
    Public Property attribute1List As New List(Of String)
    Public Property selectedAttribute01 As String
    Public Property clickedShowCavity_01Btn As Boolean = False
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        selectedAttribute01 = ComboBox1.SelectedValue.ToString()
    End Sub

    Private Sub K202_CustomSchedulingWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clickedShowCavity_01Btn = True
        Close()

    End Sub

End Class