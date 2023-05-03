Imports System.Data.SqlClient

Public Class UpstreamOrderGenarateForm
    Public Property connetionString As String

    Private Sub btn_Genarate_Click(sender As Object, e As EventArgs) Handles btn_Genarate.Click
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K202_GenarateUpstreamOrder_Sp"
            Dim param As SqlParameter

            param = New SqlParameter("@StartTime", DateTime_Start.Text)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@EndTime", DateTime_End.Text)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "1") Then
                MsgBox("Order Genarate completed",, "Preactor")
            Else
                MsgBox("Orders genarate fail",, "Preactor")
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Orders genarate error",, "error")
            ''MsgBox(ex.Message)
        Finally

        End Try
    End Sub
End Class