<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class QuantityPlanningInterface
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
        Me.GridViewMasterProduction = New System.Windows.Forms.DataGridView()
        Me.SaveBtn2 = New System.Windows.Forms.Button()
        Me.txtSKU = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnFilter = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSalesReport = New System.Windows.Forms.TextBox()
        CType(Me.GridViewMasterProduction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridViewMasterProduction
        '
        Me.GridViewMasterProduction.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridViewMasterProduction.Location = New System.Drawing.Point(10, 61)
        Me.GridViewMasterProduction.Name = "GridViewMasterProduction"
        Me.GridViewMasterProduction.Size = New System.Drawing.Size(1350, 283)
        Me.GridViewMasterProduction.TabIndex = 0
        '
        'SaveBtn2
        '
        Me.SaveBtn2.Location = New System.Drawing.Point(1215, 352)
        Me.SaveBtn2.Name = "SaveBtn2"
        Me.SaveBtn2.Size = New System.Drawing.Size(146, 28)
        Me.SaveBtn2.TabIndex = 2
        Me.SaveBtn2.Text = "Order Genarate"
        Me.SaveBtn2.UseVisualStyleBackColor = True
        Me.SaveBtn2.Visible = False
        '
        'txtSKU
        '
        Me.txtSKU.Location = New System.Drawing.Point(66, 21)
        Me.txtSKU.Name = "txtSKU"
        Me.txtSKU.Size = New System.Drawing.Size(192, 20)
        Me.txtSKU.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calisto MT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "SKU"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(739, 21)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(63, 20)
        Me.btnClear.TabIndex = 9
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(640, 22)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(61, 19)
        Me.btnFilter.TabIndex = 10
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calisto MT", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(285, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 15)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Sales Report"
        '
        'txtSalesReport
        '
        Me.txtSalesReport.Location = New System.Drawing.Point(380, 21)
        Me.txtSalesReport.Name = "txtSalesReport"
        Me.txtSalesReport.Size = New System.Drawing.Size(192, 20)
        Me.txtSalesReport.TabIndex = 7
        '
        'QuantityPlanningInterface
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1370, 411)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnFilter)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSalesReport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSKU)
        Me.Controls.Add(Me.SaveBtn2)
        Me.Controls.Add(Me.GridViewMasterProduction)
        Me.Name = "QuantityPlanningInterface"
        Me.Text = "Master Production Schedule"
        CType(Me.GridViewMasterProduction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GridViewMasterProduction As Windows.Forms.DataGridView
    Friend WithEvents SaveBtn2 As Windows.Forms.Button
    Friend WithEvents txtSKU As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnClear As Windows.Forms.Button
    Friend WithEvents btnFilter As Windows.Forms.Button
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents txtSalesReport As Windows.Forms.TextBox
End Class
