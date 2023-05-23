<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class K202_CustomSchedulingWindowSecondPart
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Ok_Btn = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(52, 27)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(479, 213)
        Me.DataGridView1.TabIndex = 0
        '
        'Ok_Btn
        '
        Me.Ok_Btn.Location = New System.Drawing.Point(456, 261)
        Me.Ok_Btn.Name = "Ok_Btn"
        Me.Ok_Btn.Size = New System.Drawing.Size(75, 23)
        Me.Ok_Btn.TabIndex = 1
        Me.Ok_Btn.Text = "Ok"
        Me.Ok_Btn.UseVisualStyleBackColor = True
        '
        'CustomSchedulingWindowSecondPart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 301)
        Me.Controls.Add(Me.Ok_Btn)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "CustomSchedulingWindowSecondPart"
        Me.Text = "CustomSchedulingWindow"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents Ok_Btn As Windows.Forms.Button
End Class
