<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Time
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
        Me.components = New System.ComponentModel.Container()
        Me.TimerClock = New System.Windows.Forms.Timer(Me.components)
        Me.TxtTime = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'TimerClock
        '
        Me.TimerClock.Enabled = True
        Me.TimerClock.Interval = 5000
        '
        'TxtTime
        '
        Me.TxtTime.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtTime.Font = New System.Drawing.Font("微軟正黑體", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TxtTime.Location = New System.Drawing.Point(12, 12)
        Me.TxtTime.Name = "TxtTime"
        Me.TxtTime.ReadOnly = True
        Me.TxtTime.Size = New System.Drawing.Size(248, 32)
        Me.TxtTime.TabIndex = 0
        Me.TxtTime.Text = "2024-03-20 18:00:00"
        '
        'Time
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(272, 56)
        Me.Controls.Add(Me.TxtTime)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Time"
        Me.Text = "Time"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TimerClock As Timer
    Friend WithEvents TxtTime As TextBox
End Class
