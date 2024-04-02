Public Class Time

    Private Sub Time_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtTime.Text = Now.ToString("yyyy-MM-dd HH:mm:ss")
        Clipboard.SetText(TxtTime.Text)
    End Sub
    Private Sub TimerClock_Tick(sender As Object, e As EventArgs) Handles TimerClock.Tick
        TxtTime.Text = Now.ToString("yyyy-MM-dd HH:mm:ss")
    End Sub
End Class