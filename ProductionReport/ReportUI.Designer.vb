<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ReportUI
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MenuBar = New System.Windows.Forms.MenuStrip()
        Me.Query = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportUI_DataGridView = New System.Windows.Forms.DataGridView()
        Me.TimerRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.AreaName = New System.Windows.Forms.ComboBox()
        Me.Btn_TEST = New System.Windows.Forms.Button()
        Me.Notice = New System.Windows.Forms.Label()
        Me.Btn_First = New System.Windows.Forms.Button()
        Me.TxtLot = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btn_RefreshStop = New System.Windows.Forms.Button()
        Me.MenuBar.SuspendLayout()
        CType(Me.ReportUI_DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuBar
        '
        Me.MenuBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Query})
        Me.MenuBar.Location = New System.Drawing.Point(0, 0)
        Me.MenuBar.Name = "MenuBar"
        Me.MenuBar.Size = New System.Drawing.Size(1584, 24)
        Me.MenuBar.TabIndex = 154
        Me.MenuBar.Text = "菜單"
        '
        'Query
        '
        Me.Query.Name = "Query"
        Me.Query.Size = New System.Drawing.Size(43, 20)
        Me.Query.Text = "查詢"
        '
        'ReportUI_DataGridView
        '
        Me.ReportUI_DataGridView.AllowUserToAddRows = False
        Me.ReportUI_DataGridView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ReportUI_DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader
        Me.ReportUI_DataGridView.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        Me.ReportUI_DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ReportUI_DataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.ReportUI_DataGridView.Location = New System.Drawing.Point(0, 56)
        Me.ReportUI_DataGridView.Name = "ReportUI_DataGridView"
        Me.ReportUI_DataGridView.RowTemplate.Height = 24
        Me.ReportUI_DataGridView.Size = New System.Drawing.Size(1584, 539)
        Me.ReportUI_DataGridView.TabIndex = 155
        '
        'TimerRefresh
        '
        Me.TimerRefresh.Interval = 300000
        '
        'AreaName
        '
        Me.AreaName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.AreaName.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AreaName.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.AreaName.FormattingEnabled = True
        Me.AreaName.Location = New System.Drawing.Point(12, 30)
        Me.AreaName.Name = "AreaName"
        Me.AreaName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.AreaName.Size = New System.Drawing.Size(167, 25)
        Me.AreaName.TabIndex = 156
        '
        'Btn_TEST
        '
        Me.Btn_TEST.Location = New System.Drawing.Point(1308, 31)
        Me.Btn_TEST.Name = "Btn_TEST"
        Me.Btn_TEST.Size = New System.Drawing.Size(75, 23)
        Me.Btn_TEST.TabIndex = 157
        Me.Btn_TEST.Text = "測試"
        Me.Btn_TEST.UseVisualStyleBackColor = True
        Me.Btn_TEST.Visible = False
        '
        'Notice
        '
        Me.Notice.AutoSize = True
        Me.Notice.Font = New System.Drawing.Font("新細明體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Notice.ForeColor = System.Drawing.Color.Red
        Me.Notice.Location = New System.Drawing.Point(526, 32)
        Me.Notice.Name = "Notice"
        Me.Notice.Size = New System.Drawing.Size(0, 19)
        Me.Notice.TabIndex = 158
        '
        'Btn_First
        '
        Me.Btn_First.Location = New System.Drawing.Point(216, 31)
        Me.Btn_First.Name = "Btn_First"
        Me.Btn_First.Size = New System.Drawing.Size(75, 23)
        Me.Btn_First.TabIndex = 159
        Me.Btn_First.Text = "首件"
        Me.Btn_First.UseVisualStyleBackColor = True
        Me.Btn_First.Visible = False
        '
        'TxtLot
        '
        Me.TxtLot.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TxtLot.Location = New System.Drawing.Point(361, 31)
        Me.TxtLot.Name = "TxtLot"
        Me.TxtLot.Size = New System.Drawing.Size(159, 25)
        Me.TxtLot.TabIndex = 160
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(310, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 20)
        Me.Label1.TabIndex = 161
        Me.Label1.Text = "批號:"
        '
        'Btn_RefreshStop
        '
        Me.Btn_RefreshStop.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Btn_RefreshStop.Location = New System.Drawing.Point(1200, 12)
        Me.Btn_RefreshStop.Name = "Btn_RefreshStop"
        Me.Btn_RefreshStop.Size = New System.Drawing.Size(102, 42)
        Me.Btn_RefreshStop.TabIndex = 162
        Me.Btn_RefreshStop.Text = "暫停刷新" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Refresh Stop)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Btn_RefreshStop.UseVisualStyleBackColor = True
        '
        'ReportUI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1584, 599)
        Me.Controls.Add(Me.Btn_RefreshStop)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtLot)
        Me.Controls.Add(Me.Btn_First)
        Me.Controls.Add(Me.Notice)
        Me.Controls.Add(Me.Btn_TEST)
        Me.Controls.Add(Me.AreaName)
        Me.Controls.Add(Me.ReportUI_DataGridView)
        Me.Controls.Add(Me.MenuBar)
        Me.Name = "ReportUI"
        Me.Text = "ReportUI [維運 : 李博軒]"
        Me.MenuBar.ResumeLayout(False)
        Me.MenuBar.PerformLayout()
        CType(Me.ReportUI_DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuBar As MenuStrip
    Friend WithEvents Query As ToolStripMenuItem
    Friend WithEvents ReportUI_DataGridView As DataGridView
    Friend WithEvents TimerRefresh As Timer
    Friend WithEvents AreaName As ComboBox
    Friend WithEvents Btn_TEST As Button
    Friend WithEvents Notice As Label
    Friend WithEvents Btn_First As Button
    Friend WithEvents TxtLot As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Btn_RefreshStop As Button
End Class
