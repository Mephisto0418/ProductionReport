<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SAP_PnlKeyIn
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Btn_Confirm = New System.Windows.Forms.Button()
        Me.dgv_PnlKeyIn = New System.Windows.Forms.DataGridView()
        Me.序號 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.基板編號 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgv_PnlKeyIn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_Confirm
        '
        Me.Btn_Confirm.Font = New System.Drawing.Font("微軟正黑體", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Btn_Confirm.Location = New System.Drawing.Point(197, 12)
        Me.Btn_Confirm.Name = "Btn_Confirm"
        Me.Btn_Confirm.Size = New System.Drawing.Size(108, 36)
        Me.Btn_Confirm.TabIndex = 0
        Me.Btn_Confirm.Text = "確定"
        Me.Btn_Confirm.UseVisualStyleBackColor = True
        '
        'dgv_PnlKeyIn
        '
        Me.dgv_PnlKeyIn.AllowUserToAddRows = False
        Me.dgv_PnlKeyIn.AllowUserToDeleteRows = False
        Me.dgv_PnlKeyIn.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_PnlKeyIn.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.序號, Me.基板編號})
        Me.dgv_PnlKeyIn.Location = New System.Drawing.Point(12, 54)
        Me.dgv_PnlKeyIn.Name = "dgv_PnlKeyIn"
        Me.dgv_PnlKeyIn.RowTemplate.Height = 24
        Me.dgv_PnlKeyIn.Size = New System.Drawing.Size(293, 419)
        Me.dgv_PnlKeyIn.TabIndex = 1
        '
        '序號
        '
        Me.序號.HeaderText = "序號"
        Me.序號.Name = "序號"
        Me.序號.ReadOnly = True
        '
        '基板編號
        '
        Me.基板編號.HeaderText = "基板編號"
        Me.基板編號.Name = "基板編號"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Location = New System.Drawing.Point(14, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "請依順序輸入基板編號 !"
        '
        'SAP_PnlKeyIn
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 485)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgv_PnlKeyIn)
        Me.Controls.Add(Me.Btn_Confirm)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "SAP_PnlKeyIn"
        Me.Text = "SAP_PnlKeyIn"
        Me.TopMost = True
        CType(Me.dgv_PnlKeyIn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Btn_Confirm As Button
    Friend WithEvents dgv_PnlKeyIn As DataGridView
    Friend WithEvents 序號 As DataGridViewTextBoxColumn
    Friend WithEvents 基板編號 As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
End Class
