Public Class mainForm

    Dim WithEvents contButton As New continueButton

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = Windows.Forms.FormWindowState.Maximized
        FormBorderStyle = Windows.Forms.FormBorderStyle.None
        BackColor = Color.White
        Controls.Add(contButton)

    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles contButton.Click

        Call mainTest()

    End Sub





End Class
