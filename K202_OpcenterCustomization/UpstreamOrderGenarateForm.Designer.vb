<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UpstreamOrderGenarateForm
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
        Me.btn_Genarate = New System.Windows.Forms.Button()
        Me.DateTime_End = New System.Windows.Forms.DateTimePicker()
        Me.DateTime_Start = New System.Windows.Forms.DateTimePicker()
        Me.Label_EndDate = New System.Windows.Forms.Label()
        Me.Label_StartDate = New System.Windows.Forms.Label()
        Me.Label_UpstreamOrderGenarate = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btn_Genarate
        '
        Me.btn_Genarate.Location = New System.Drawing.Point(200, 114)
        Me.btn_Genarate.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_Genarate.Name = "btn_Genarate"
        Me.btn_Genarate.Size = New System.Drawing.Size(107, 19)
        Me.btn_Genarate.TabIndex = 11
        Me.btn_Genarate.Text = "Genarate"
        Me.btn_Genarate.UseVisualStyleBackColor = True
        '
        'DateTime_End
        '
        Me.DateTime_End.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTime_End.Location = New System.Drawing.Point(351, 75)
        Me.DateTime_End.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTime_End.Name = "DateTime_End"
        Me.DateTime_End.Size = New System.Drawing.Size(135, 20)
        Me.DateTime_End.TabIndex = 10
        '
        'DateTime_Start
        '
        Me.DateTime_Start.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTime_Start.Location = New System.Drawing.Point(88, 75)
        Me.DateTime_Start.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTime_Start.Name = "DateTime_Start"
        Me.DateTime_Start.Size = New System.Drawing.Size(121, 20)
        Me.DateTime_Start.TabIndex = 9
        '
        'Label_EndDate
        '
        Me.Label_EndDate.AutoSize = True
        Me.Label_EndDate.Location = New System.Drawing.Point(296, 75)
        Me.Label_EndDate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label_EndDate.Name = "Label_EndDate"
        Me.Label_EndDate.Size = New System.Drawing.Size(52, 13)
        Me.Label_EndDate.TabIndex = 8
        Me.Label_EndDate.Text = "End Date"
        '
        'Label_StartDate
        '
        Me.Label_StartDate.AutoSize = True
        Me.Label_StartDate.Location = New System.Drawing.Point(29, 75)
        Me.Label_StartDate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label_StartDate.Name = "Label_StartDate"
        Me.Label_StartDate.Size = New System.Drawing.Size(55, 13)
        Me.Label_StartDate.TabIndex = 7
        Me.Label_StartDate.Text = "Start Date"
        '
        'Label_UpstreamOrderGenarate
        '
        Me.Label_UpstreamOrderGenarate.AutoSize = True
        Me.Label_UpstreamOrderGenarate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_UpstreamOrderGenarate.Location = New System.Drawing.Point(152, 22)
        Me.Label_UpstreamOrderGenarate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label_UpstreamOrderGenarate.Name = "Label_UpstreamOrderGenarate"
        Me.Label_UpstreamOrderGenarate.Size = New System.Drawing.Size(226, 24)
        Me.Label_UpstreamOrderGenarate.TabIndex = 6
        Me.Label_UpstreamOrderGenarate.Text = "Upstream Order Genarate"
        '
        'UpstreamOrderGenarateForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(541, 184)
        Me.Controls.Add(Me.btn_Genarate)
        Me.Controls.Add(Me.DateTime_End)
        Me.Controls.Add(Me.DateTime_Start)
        Me.Controls.Add(Me.Label_EndDate)
        Me.Controls.Add(Me.Label_StartDate)
        Me.Controls.Add(Me.Label_UpstreamOrderGenarate)
        Me.Name = "UpstreamOrderGenarateForm"
        Me.Text = "UpstreamOrderGenarateForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_Genarate As Windows.Forms.Button
    Friend WithEvents DateTime_End As Windows.Forms.DateTimePicker
    Friend WithEvents DateTime_Start As Windows.Forms.DateTimePicker
    Friend WithEvents Label_EndDate As Windows.Forms.Label
    Friend WithEvents Label_StartDate As Windows.Forms.Label
    Friend WithEvents Label_UpstreamOrderGenarate As Windows.Forms.Label
End Class
