Public Class subjectForm
    Inherits Form

    Private WithEvents contButton As New continueButton(Txt:="Bestätigen")
    Private WithEvents subjBox As New TextBox
    Private ReadOnly subjPanel As New subjectPanel(subjBox)

    Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

        WindowState = FormWindowState.Normal
        FormBorderStyle = FormBorderStyle.None

        Size = New Point(250, 200)
        BackColor = Color.Gray
        Top = (3 * screenHeight / 4) - (Height / 2)
        Left = (screenWidth / 2) - (Width / 2)

        Controls.Add(contButton)
        xCenter(contButton)

        Controls.Add(subjPanel)
        xCenter(subjPanel, 0.4)

        subjBox.Select()

    End Sub

    Public condN As Integer
    Public subjN As Integer
    Public keyAss As String

    Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

        If subjBox.Text = "" OrElse Val(subjBox.Text) < 1 OrElse Val(subjBox.Text) > 500 Then
            MsgBox("Bitte geben Sie eine korrekte VPNr ein!", MsgBoxStyle.Critical, Title:="Fehler!")
            Exit Sub
        Else
            subjN = CInt(subjBox.Text)
            condN = setCond(subjN)
            Select Case condN
                Case 0
                    keyAss = "Apos"
                Case 1
                    keyAss = "Aneg"
            End Select
            mainForm.dataFrame("Subject") = subjN.ToString
            mainForm.dataFrame("Key") = keyAss
        End If

        mainForm.contButton.PerformClick()
        Close()
    End Sub

    Private Sub confirmEnter(sender As Object, e As KeyEventArgs) Handles subjBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            contButton.PerformClick()
        End If
    End Sub

    Private Sub suppressNonNumeric(sender As Object, e As KeyPressEventArgs) Handles subjBox.KeyPress
        If Not IsNumeric(e.KeyChar) And e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

End Class