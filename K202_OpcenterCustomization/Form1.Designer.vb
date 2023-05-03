<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.SalesOrderComboBox = New System.Windows.Forms.ComboBox()
        Me.SKUComboBox = New System.Windows.Forms.ComboBox()
        Me.DataGridView = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SalesOrderComboBox
        '
        Me.SalesOrderComboBox.FormattingEnabled = True
        Me.SalesOrderComboBox.Location = New System.Drawing.Point(129, 70)
        Me.SalesOrderComboBox.Name = "SalesOrderComboBox"
        Me.SalesOrderComboBox.Size = New System.Drawing.Size(121, 21)
        Me.SalesOrderComboBox.TabIndex = 0
        '
        'SKUComboBox
        '
        Me.SKUComboBox.FormattingEnabled = True
        Me.SKUComboBox.Location = New System.Drawing.Point(405, 69)
        Me.SKUComboBox.Name = "SKUComboBox"
        Me.SKUComboBox.Size = New System.Drawing.Size(121, 21)
        Me.SKUComboBox.TabIndex = 1
        '
        'DataGridView
        '
        Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView.Location = New System.Drawing.Point(39, 143)
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.Size = New System.Drawing.Size(851, 273)
        Me.DataGridView.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(926, 450)
        Me.Controls.Add(Me.DataGridView)
        Me.Controls.Add(Me.SKUComboBox)
        Me.Controls.Add(Me.SalesOrderComboBox)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SalesOrderComboBox As Windows.Forms.ComboBox
    Friend WithEvents SKUComboBox As Windows.Forms.ComboBox
    Friend WithEvents DataGridView As Windows.Forms.DataGridView
End Class
