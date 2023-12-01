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
        Me.cboAreaName = New System.Windows.Forms.ComboBox()
        Me.Btn_TEST = New System.Windows.Forms.Button()
        Me.Notice = New System.Windows.Forms.Label()
        Me.Btn_First = New System.Windows.Forms.Button()
        Me.TxtLot = New System.Windows.Forms.TextBox()
        Me.lblLot = New System.Windows.Forms.Label()
        Me.Btn_RefreshStop = New System.Windows.Forms.Button()
        Me.lblMachine = New System.Windows.Forms.Label()
        Me.cboMachine = New System.Windows.Forms.ComboBox()
        Me.lblEndTime = New System.Windows.Forms.Label()
        Me.dtpEndTime = New System.Windows.Forms.DateTimePicker()
        Me.dtpStartTime = New System.Windows.Forms.DateTimePicker()
        Me.lblStartTime = New System.Windows.Forms.Label()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblRemark = New System.Windows.Forms.Label()
        Me.txtRemark = New System.Windows.Forms.TextBox()
        Me.grpRemark = New System.Windows.Forms.GroupBox()
        Me.btnRemarkUpload = New System.Windows.Forms.Button()
        Me.txtMachineState = New System.Windows.Forms.TextBox()
        Me.lblMachineState = New System.Windows.Forms.Label()
        Me.lblMachine_Eng = New System.Windows.Forms.Label()
        Me.lblType_Eng = New System.Windows.Forms.Label()
        Me.lblRemark_Eng = New System.Windows.Forms.Label()
        Me.lblMachineState_Eng = New System.Windows.Forms.Label()
        Me.lblStartTime_Eng = New System.Windows.Forms.Label()
        Me.lblEndTime_Eng = New System.Windows.Forms.Label()
        Me.lblRemarkUpload_Eng = New System.Windows.Forms.Label()
        Me.lblLot_Eng = New System.Windows.Forms.Label()
        Me.MenuBar.SuspendLayout()
        CType(Me.ReportUI_DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpRemark.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuBar
        '
        Me.MenuBar.BackColor = System.Drawing.Color.Transparent
        Me.MenuBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Query})
        Me.MenuBar.Location = New System.Drawing.Point(0, 0)
        Me.MenuBar.Name = "MenuBar"
        Me.MenuBar.Size = New System.Drawing.Size(1584, 25)
        Me.MenuBar.TabIndex = 154
        Me.MenuBar.Text = "菜單"
        '
        'Query
        '
        Me.Query.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Query.Name = "Query"
        Me.Query.Size = New System.Drawing.Size(46, 21)
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
        Me.ReportUI_DataGridView.Location = New System.Drawing.Point(0, 97)
        Me.ReportUI_DataGridView.Name = "ReportUI_DataGridView"
        Me.ReportUI_DataGridView.RowTemplate.Height = 24
        Me.ReportUI_DataGridView.Size = New System.Drawing.Size(1584, 498)
        Me.ReportUI_DataGridView.TabIndex = 155
        '
        'TimerRefresh
        '
        Me.TimerRefresh.Interval = 300000
        '
        'cboAreaName
        '
        Me.cboAreaName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAreaName.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cboAreaName.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cboAreaName.FormattingEnabled = True
        Me.cboAreaName.Location = New System.Drawing.Point(12, 68)
        Me.cboAreaName.Name = "cboAreaName"
        Me.cboAreaName.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.cboAreaName.Size = New System.Drawing.Size(167, 25)
        Me.cboAreaName.TabIndex = 156
        '
        'Btn_TEST
        '
        Me.Btn_TEST.Location = New System.Drawing.Point(1507, 28)
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
        Me.Notice.Location = New System.Drawing.Point(310, 48)
        Me.Notice.Name = "Notice"
        Me.Notice.Size = New System.Drawing.Size(0, 19)
        Me.Notice.TabIndex = 158
        '
        'Btn_First
        '
        Me.Btn_First.Location = New System.Drawing.Point(216, 69)
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
        Me.TxtLot.Location = New System.Drawing.Point(361, 69)
        Me.TxtLot.Name = "TxtLot"
        Me.TxtLot.Size = New System.Drawing.Size(159, 25)
        Me.TxtLot.TabIndex = 160
        '
        'lblLot
        '
        Me.lblLot.AutoSize = True
        Me.lblLot.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblLot.Location = New System.Drawing.Point(310, 65)
        Me.lblLot.Name = "lblLot"
        Me.lblLot.Size = New System.Drawing.Size(45, 20)
        Me.lblLot.TabIndex = 161
        Me.lblLot.Text = "批號:"
        '
        'Btn_RefreshStop
        '
        Me.Btn_RefreshStop.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Btn_RefreshStop.Location = New System.Drawing.Point(1480, 50)
        Me.Btn_RefreshStop.Name = "Btn_RefreshStop"
        Me.Btn_RefreshStop.Size = New System.Drawing.Size(102, 42)
        Me.Btn_RefreshStop.TabIndex = 162
        Me.Btn_RefreshStop.Text = "暫停刷新" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Refresh Stop)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Btn_RefreshStop.UseVisualStyleBackColor = True
        '
        'lblMachine
        '
        Me.lblMachine.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblMachine.Location = New System.Drawing.Point(3, 13)
        Me.lblMachine.Name = "lblMachine"
        Me.lblMachine.Size = New System.Drawing.Size(77, 20)
        Me.lblMachine.TabIndex = 164
        Me.lblMachine.Text = "機台:"
        Me.lblMachine.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboMachine
        '
        Me.cboMachine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMachine.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cboMachine.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cboMachine.FormattingEnabled = True
        Me.cboMachine.Location = New System.Drawing.Point(86, 10)
        Me.cboMachine.Name = "cboMachine"
        Me.cboMachine.Size = New System.Drawing.Size(150, 24)
        Me.cboMachine.TabIndex = 165
        '
        'lblEndTime
        '
        Me.lblEndTime.AutoSize = True
        Me.lblEndTime.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblEndTime.Location = New System.Drawing.Point(245, 50)
        Me.lblEndTime.Name = "lblEndTime"
        Me.lblEndTime.Size = New System.Drawing.Size(77, 20)
        Me.lblEndTime.TabIndex = 166
        Me.lblEndTime.Text = "結束時間:"
        '
        'dtpEndTime
        '
        Me.dtpEndTime.CalendarFont = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.dtpEndTime.CustomFormat = "yyyy/MM/dd HH:mm"
        Me.dtpEndTime.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.dtpEndTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEndTime.Location = New System.Drawing.Point(328, 50)
        Me.dtpEndTime.Name = "dtpEndTime"
        Me.dtpEndTime.Size = New System.Drawing.Size(150, 23)
        Me.dtpEndTime.TabIndex = 167
        '
        'dtpStartTime
        '
        Me.dtpStartTime.CalendarFont = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.dtpStartTime.CustomFormat = "yyyy/MM/dd HH:mm"
        Me.dtpStartTime.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.dtpStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartTime.Location = New System.Drawing.Point(86, 50)
        Me.dtpStartTime.Name = "dtpStartTime"
        Me.dtpStartTime.Size = New System.Drawing.Size(150, 23)
        Me.dtpStartTime.TabIndex = 169
        '
        'lblStartTime
        '
        Me.lblStartTime.AutoSize = True
        Me.lblStartTime.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblStartTime.Location = New System.Drawing.Point(3, 50)
        Me.lblStartTime.Name = "lblStartTime"
        Me.lblStartTime.Size = New System.Drawing.Size(77, 20)
        Me.lblStartTime.TabIndex = 168
        Me.lblStartTime.Text = "開始時間:"
        '
        'cboType
        '
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cboType.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cboType.FormattingEnabled = True
        Me.cboType.Location = New System.Drawing.Point(328, 10)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(150, 24)
        Me.cboType.TabIndex = 171
        '
        'lblType
        '
        Me.lblType.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblType.Location = New System.Drawing.Point(245, 13)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(77, 20)
        Me.lblType.TabIndex = 170
        Me.lblType.Text = "分類:"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRemark
        '
        Me.lblRemark.AutoSize = True
        Me.lblRemark.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblRemark.Location = New System.Drawing.Point(495, 13)
        Me.lblRemark.Name = "lblRemark"
        Me.lblRemark.Size = New System.Drawing.Size(45, 20)
        Me.lblRemark.TabIndex = 173
        Me.lblRemark.Text = "備註:"
        '
        'txtRemark
        '
        Me.txtRemark.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.txtRemark.Location = New System.Drawing.Point(546, 10)
        Me.txtRemark.Name = "txtRemark"
        Me.txtRemark.Size = New System.Drawing.Size(308, 25)
        Me.txtRemark.TabIndex = 172
        '
        'grpRemark
        '
        Me.grpRemark.BackColor = System.Drawing.Color.Transparent
        Me.grpRemark.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.grpRemark.Controls.Add(Me.btnRemarkUpload)
        Me.grpRemark.Controls.Add(Me.txtMachineState)
        Me.grpRemark.Controls.Add(Me.lblMachineState)
        Me.grpRemark.Controls.Add(Me.lblRemark)
        Me.grpRemark.Controls.Add(Me.cboMachine)
        Me.grpRemark.Controls.Add(Me.txtRemark)
        Me.grpRemark.Controls.Add(Me.lblMachine)
        Me.grpRemark.Controls.Add(Me.cboType)
        Me.grpRemark.Controls.Add(Me.lblEndTime)
        Me.grpRemark.Controls.Add(Me.lblType)
        Me.grpRemark.Controls.Add(Me.dtpEndTime)
        Me.grpRemark.Controls.Add(Me.dtpStartTime)
        Me.grpRemark.Controls.Add(Me.lblStartTime)
        Me.grpRemark.Controls.Add(Me.lblMachine_Eng)
        Me.grpRemark.Controls.Add(Me.lblType_Eng)
        Me.grpRemark.Controls.Add(Me.lblRemark_Eng)
        Me.grpRemark.Controls.Add(Me.lblMachineState_Eng)
        Me.grpRemark.Controls.Add(Me.lblStartTime_Eng)
        Me.grpRemark.Controls.Add(Me.lblEndTime_Eng)
        Me.grpRemark.Controls.Add(Me.lblRemarkUpload_Eng)
        Me.grpRemark.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.grpRemark.Location = New System.Drawing.Point(583, 12)
        Me.grpRemark.Name = "grpRemark"
        Me.grpRemark.Size = New System.Drawing.Size(860, 83)
        Me.grpRemark.TabIndex = 174
        Me.grpRemark.TabStop = False
        Me.grpRemark.Text = "批間作業紀錄"
        '
        'btnRemarkUpload
        '
        Me.btnRemarkUpload.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnRemarkUpload.Location = New System.Drawing.Point(714, 48)
        Me.btnRemarkUpload.Name = "btnRemarkUpload"
        Me.btnRemarkUpload.Size = New System.Drawing.Size(75, 23)
        Me.btnRemarkUpload.TabIndex = 175
        Me.btnRemarkUpload.Text = "備註上傳"
        Me.btnRemarkUpload.UseVisualStyleBackColor = True
        '
        'txtMachineState
        '
        Me.txtMachineState.BackColor = System.Drawing.Color.Gold
        Me.txtMachineState.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMachineState.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.txtMachineState.Location = New System.Drawing.Point(546, 50)
        Me.txtMachineState.Name = "txtMachineState"
        Me.txtMachineState.ReadOnly = True
        Me.txtMachineState.Size = New System.Drawing.Size(108, 18)
        Me.txtMachineState.TabIndex = 175
        Me.txtMachineState.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblMachineState
        '
        Me.lblMachineState.AutoSize = True
        Me.lblMachineState.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblMachineState.Location = New System.Drawing.Point(495, 50)
        Me.lblMachineState.Name = "lblMachineState"
        Me.lblMachineState.Size = New System.Drawing.Size(45, 20)
        Me.lblMachineState.TabIndex = 174
        Me.lblMachineState.Text = "機況:"
        '
        'lblMachine_Eng
        '
        Me.lblMachine_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblMachine_Eng.Location = New System.Drawing.Point(10, 29)
        Me.lblMachine_Eng.Name = "lblMachine_Eng"
        Me.lblMachine_Eng.Size = New System.Drawing.Size(77, 20)
        Me.lblMachine_Eng.TabIndex = 176
        Me.lblMachine_Eng.Text = "(Machine)"
        Me.lblMachine_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblType_Eng
        '
        Me.lblType_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblType_Eng.Location = New System.Drawing.Point(243, 29)
        Me.lblType_Eng.Name = "lblType_Eng"
        Me.lblType_Eng.Size = New System.Drawing.Size(77, 20)
        Me.lblType_Eng.TabIndex = 177
        Me.lblType_Eng.Text = "(Type)"
        Me.lblType_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRemark_Eng
        '
        Me.lblRemark_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblRemark_Eng.Location = New System.Drawing.Point(467, 29)
        Me.lblRemark_Eng.Name = "lblRemark_Eng"
        Me.lblRemark_Eng.Size = New System.Drawing.Size(77, 20)
        Me.lblRemark_Eng.TabIndex = 178
        Me.lblRemark_Eng.Text = "(Remark)"
        Me.lblRemark_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblMachineState_Eng
        '
        Me.lblMachineState_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblMachineState_Eng.Location = New System.Drawing.Point(476, 65)
        Me.lblMachineState_Eng.Name = "lblMachineState_Eng"
        Me.lblMachineState_Eng.Size = New System.Drawing.Size(87, 20)
        Me.lblMachineState_Eng.TabIndex = 177
        Me.lblMachineState_Eng.Text = "(MachineState)"
        Me.lblMachineState_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStartTime_Eng
        '
        Me.lblStartTime_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblStartTime_Eng.Location = New System.Drawing.Point(-4, 65)
        Me.lblStartTime_Eng.Name = "lblStartTime_Eng"
        Me.lblStartTime_Eng.Size = New System.Drawing.Size(77, 20)
        Me.lblStartTime_Eng.TabIndex = 179
        Me.lblStartTime_Eng.Text = "(StartTime)"
        Me.lblStartTime_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEndTime_Eng
        '
        Me.lblEndTime_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblEndTime_Eng.Location = New System.Drawing.Point(236, 65)
        Me.lblEndTime_Eng.Name = "lblEndTime_Eng"
        Me.lblEndTime_Eng.Size = New System.Drawing.Size(77, 20)
        Me.lblEndTime_Eng.TabIndex = 180
        Me.lblEndTime_Eng.Text = "(EndTime)"
        Me.lblEndTime_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRemarkUpload_Eng
        '
        Me.lblRemarkUpload_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblRemarkUpload_Eng.Location = New System.Drawing.Point(704, 65)
        Me.lblRemarkUpload_Eng.Name = "lblRemarkUpload_Eng"
        Me.lblRemarkUpload_Eng.Size = New System.Drawing.Size(98, 20)
        Me.lblRemarkUpload_Eng.TabIndex = 181
        Me.lblRemarkUpload_Eng.Text = "(Remark Upload)"
        Me.lblRemarkUpload_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLot_Eng
        '
        Me.lblLot_Eng.AutoSize = True
        Me.lblLot_Eng.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblLot_Eng.Location = New System.Drawing.Point(307, 83)
        Me.lblLot_Eng.Name = "lblLot_Eng"
        Me.lblLot_Eng.Size = New System.Drawing.Size(51, 15)
        Me.lblLot_Eng.TabIndex = 182
        Me.lblLot_Eng.Text = "(LotNo.)"
        Me.lblLot_Eng.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ReportUI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1584, 599)
        Me.Controls.Add(Me.ReportUI_DataGridView)
        Me.Controls.Add(Me.grpRemark)
        Me.Controls.Add(Me.MenuBar)
        Me.Controls.Add(Me.Btn_RefreshStop)
        Me.Controls.Add(Me.lblLot)
        Me.Controls.Add(Me.TxtLot)
        Me.Controls.Add(Me.Btn_First)
        Me.Controls.Add(Me.Notice)
        Me.Controls.Add(Me.Btn_TEST)
        Me.Controls.Add(Me.cboAreaName)
        Me.Controls.Add(Me.lblLot_Eng)
        Me.Name = "ReportUI"
        Me.Text = "ReportUI [維運 : 李博軒]"
        Me.MenuBar.ResumeLayout(False)
        Me.MenuBar.PerformLayout()
        CType(Me.ReportUI_DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpRemark.ResumeLayout(False)
        Me.grpRemark.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuBar As MenuStrip
    Friend WithEvents Query As ToolStripMenuItem
    Friend WithEvents ReportUI_DataGridView As DataGridView
    Friend WithEvents TimerRefresh As Timer
    Friend WithEvents cboAreaName As ComboBox
    Friend WithEvents Btn_TEST As Button
    Friend WithEvents Notice As Label
    Friend WithEvents Btn_First As Button
    Friend WithEvents TxtLot As TextBox
    Friend WithEvents lblLot As Label
    Friend WithEvents Btn_RefreshStop As Button
    Friend WithEvents lblMachine As Label
    Friend WithEvents cboMachine As ComboBox
    Friend WithEvents lblEndTime As Label
    Friend WithEvents dtpEndTime As DateTimePicker
    Friend WithEvents dtpStartTime As DateTimePicker
    Friend WithEvents lblStartTime As Label
    Friend WithEvents cboType As ComboBox
    Friend WithEvents lblType As Label
    Friend WithEvents lblRemark As Label
    Friend WithEvents txtRemark As TextBox
    Friend WithEvents grpRemark As GroupBox
    Friend WithEvents btnRemarkUpload As Button
    Friend WithEvents txtMachineState As TextBox
    Friend WithEvents lblMachineState As Label
    Friend WithEvents lblRemark_Eng As Label
    Friend WithEvents lblMachine_Eng As Label
    Friend WithEvents lblType_Eng As Label
    Friend WithEvents lblMachineState_Eng As Label
    Friend WithEvents lblStartTime_Eng As Label
    Friend WithEvents lblEndTime_Eng As Label
    Friend WithEvents lblRemarkUpload_Eng As Label
    Friend WithEvents lblLot_Eng As Label
End Class
