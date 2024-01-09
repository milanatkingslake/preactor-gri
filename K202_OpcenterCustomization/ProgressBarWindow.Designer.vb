<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ProgressBarWindow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.BackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(12, 12)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(507, 23)
        Me.ProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.ProgressBar.TabIndex = 0
        '
        'BackgroundWorker
        '
        Me.BackgroundWorker.WorkerReportsProgress = True
        Me.BackgroundWorker.WorkerSupportsCancellation = True
        '
        'ProgressBarWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(530, 49)
        Me.ControlBox = False
        Me.Controls.Add(Me.ProgressBar)
        Me.Name = "ProgressBarWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "System In Progress"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ProgressBar As Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker As ComponentModel.BackgroundWorker
End Class
