<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class K202_JobSplitDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OrderNo = New System.Windows.Forms.Label()
        Me.Quanty = New System.Windows.Forms.Label()
        Me.JobOrderNoTxt = New System.Windows.Forms.TextBox()
        Me.JobQtyTxt = New System.Windows.Forms.TextBox()
        Me.OkBtn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OrderNo
        '
        Me.OrderNo.AutoSize = True
        Me.OrderNo.Location = New System.Drawing.Point(105, 25)
        Me.OrderNo.Name = "OrderNo"
        Me.OrderNo.Size = New System.Drawing.Size(50, 13)
        Me.OrderNo.TabIndex = 0
        Me.OrderNo.Text = "Order No"
        '
        'Quanty
        '
        Me.Quanty.AutoSize = True
        Me.Quanty.Location = New System.Drawing.Point(105, 53)
        Me.Quanty.Name = "Quanty"
        Me.Quanty.Size = New System.Drawing.Size(41, 13)
        Me.Quanty.TabIndex = 1
        Me.Quanty.Text = "Quanty"
        '
        'JobOrderNoTxt
        '
        Me.JobOrderNoTxt.Location = New System.Drawing.Point(174, 25)
        Me.JobOrderNoTxt.Name = "JobOrderNoTxt"
        Me.JobOrderNoTxt.ReadOnly = True
        Me.JobOrderNoTxt.Size = New System.Drawing.Size(159, 20)
        Me.JobOrderNoTxt.TabIndex = 2
        '
        'JobQtyTxt
        '
        Me.JobQtyTxt.Location = New System.Drawing.Point(174, 50)
        Me.JobQtyTxt.Name = "JobQtyTxt"
        Me.JobQtyTxt.Size = New System.Drawing.Size(100, 20)
        Me.JobQtyTxt.TabIndex = 3
        '
        'OkBtn
        '
        Me.OkBtn.Location = New System.Drawing.Point(128, 84)
        Me.OkBtn.Name = "OkBtn"
        Me.OkBtn.Size = New System.Drawing.Size(75, 23)
        Me.OkBtn.TabIndex = 5
        Me.OkBtn.Text = "Ok"
        Me.OkBtn.UseVisualStyleBackColor = True
        '
        'K202_JobSplitDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(382, 128)
        Me.Controls.Add(Me.OkBtn)
        Me.Controls.Add(Me.JobQtyTxt)
        Me.Controls.Add(Me.JobOrderNoTxt)
        Me.Controls.Add(Me.Quanty)
        Me.Controls.Add(Me.OrderNo)
        Me.Name = "K202_JobSplitDetails"
        Me.Text = "K202_JobSplitDetails"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents OrderNo As Windows.Forms.Label
    Friend WithEvents Quanty As Windows.Forms.Label
    Friend WithEvents JobOrderNoTxt As Windows.Forms.TextBox
    Friend WithEvents JobQtyTxt As Windows.Forms.TextBox
    Friend WithEvents OkBtn As Windows.Forms.Button
End Class
