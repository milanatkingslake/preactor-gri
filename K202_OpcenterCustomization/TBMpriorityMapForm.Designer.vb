<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TBMpriorityMapForm
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
        Me.ComboBoxRule = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridViewMapResourceGroup = New System.Windows.Forms.DataGridView()
        Me.DataGridViewAllResourceGroup = New System.Windows.Forms.DataGridView()
        Me.btnMap = New System.Windows.Forms.Button()
        Me.btnUnMap = New System.Windows.Forms.Button()
        Me.Save = New System.Windows.Forms.Button()
        Me.lblRule = New System.Windows.Forms.Label()
        CType(Me.DataGridViewMapResourceGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewAllResourceGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBoxRule
        '
        Me.ComboBoxRule.FormattingEnabled = True
        Me.ComboBoxRule.Items.AddRange(New Object() {"2P2O", "2P3O", "3P2O"})
        Me.ComboBoxRule.Location = New System.Drawing.Point(558, 55)
        Me.ComboBoxRule.Name = "ComboBoxRule"
        Me.ComboBoxRule.Size = New System.Drawing.Size(121, 28)
        Me.ComboBoxRule.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(554, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "TBM Priority Map"
        '
        'DataGridViewMapResourceGroup
        '
        Me.DataGridViewMapResourceGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewMapResourceGroup.Location = New System.Drawing.Point(753, 113)
        Me.DataGridViewMapResourceGroup.MaximumSize = New System.Drawing.Size(380, 400)
        Me.DataGridViewMapResourceGroup.MinimumSize = New System.Drawing.Size(380, 400)
        Me.DataGridViewMapResourceGroup.Name = "DataGridViewMapResourceGroup"
        Me.DataGridViewMapResourceGroup.RowHeadersWidth = 62
        Me.DataGridViewMapResourceGroup.RowTemplate.Height = 28
        Me.DataGridViewMapResourceGroup.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridViewMapResourceGroup.Size = New System.Drawing.Size(380, 400)
        Me.DataGridViewMapResourceGroup.TabIndex = 2
        '
        'DataGridViewAllResourceGroup
        '
        Me.DataGridViewAllResourceGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewAllResourceGroup.Location = New System.Drawing.Point(27, 113)
        Me.DataGridViewAllResourceGroup.MaximumSize = New System.Drawing.Size(380, 400)
        Me.DataGridViewAllResourceGroup.MinimumSize = New System.Drawing.Size(380, 400)
        Me.DataGridViewAllResourceGroup.Name = "DataGridViewAllResourceGroup"
        Me.DataGridViewAllResourceGroup.RowHeadersWidth = 62
        Me.DataGridViewAllResourceGroup.RowTemplate.Height = 28
        Me.DataGridViewAllResourceGroup.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridViewAllResourceGroup.Size = New System.Drawing.Size(380, 400)
        Me.DataGridViewAllResourceGroup.TabIndex = 3
        '
        'btnMap
        '
        Me.btnMap.Location = New System.Drawing.Point(532, 246)
        Me.btnMap.Name = "btnMap"
        Me.btnMap.Size = New System.Drawing.Size(108, 34)
        Me.btnMap.TabIndex = 4
        Me.btnMap.Text = ">>"
        Me.btnMap.UseVisualStyleBackColor = True
        '
        'btnUnMap
        '
        Me.btnUnMap.Location = New System.Drawing.Point(533, 310)
        Me.btnUnMap.Name = "btnUnMap"
        Me.btnUnMap.Size = New System.Drawing.Size(110, 36)
        Me.btnUnMap.TabIndex = 5
        Me.btnUnMap.Text = "<<"
        Me.btnUnMap.UseVisualStyleBackColor = True
        '
        'Save
        '
        Me.Save.Location = New System.Drawing.Point(870, 522)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(140, 43)
        Me.Save.TabIndex = 6
        Me.Save.Text = "Save"
        Me.Save.UseVisualStyleBackColor = True
        '
        'lblRule
        '
        Me.lblRule.AutoSize = True
        Me.lblRule.Location = New System.Drawing.Point(510, 58)
        Me.lblRule.Name = "lblRule"
        Me.lblRule.Size = New System.Drawing.Size(42, 20)
        Me.lblRule.TabIndex = 7
        Me.lblRule.Text = "Rule"
        '
        'TBMpriorityMapForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1188, 574)
        Me.Controls.Add(Me.lblRule)
        Me.Controls.Add(Me.Save)
        Me.Controls.Add(Me.btnUnMap)
        Me.Controls.Add(Me.btnMap)
        Me.Controls.Add(Me.DataGridViewAllResourceGroup)
        Me.Controls.Add(Me.DataGridViewMapResourceGroup)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxRule)
        Me.Name = "TBMpriorityMapForm"
        Me.Text = "TBMpriorityMapForm"
        CType(Me.DataGridViewMapResourceGroup, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewAllResourceGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBoxRule As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents DataGridViewMapResourceGroup As Windows.Forms.DataGridView
    Friend WithEvents DataGridViewAllResourceGroup As Windows.Forms.DataGridView
    Friend WithEvents btnMap As Windows.Forms.Button
    Friend WithEvents btnUnMap As Windows.Forms.Button
    Friend WithEvents Save As Windows.Forms.Button
    Friend WithEvents lblRule As Windows.Forms.Label
End Class
