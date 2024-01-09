Imports System.ComponentModel
Imports System.Threading

Public Class ProgressBarWindow

    Public maxProgress As Integer
    Private Sub ProgressBarWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProgressBar.Value = 0
        ProgressBar.Maximum = maxProgress
        BackgroundWorker.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker.RunWorkerCompleted

    End Sub

    Private Sub BackgroundWorker_ProgressChanged(sender As Object, e As ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker.ProgressChanged

        ProgressBar.Value += 1
    End Sub

    Private Sub BackgroundWorker_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork
        Dim i As Integer

        For i = 1 To maxProgress
            Thread.Sleep(1000)
            BackgroundWorker.ReportProgress(0)
        Next i


    End Sub
End Class