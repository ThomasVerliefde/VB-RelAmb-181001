Public Class subjectForm
    Inherits Form

    Private WithEvents contButton As New continueButton(Txt:="Bestätigen")
    Private WithEvents subjBox As New TextBox
	Private ReadOnly subjPanel As New labelledBox(Me.subjBox)

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
		Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

		Me.WindowState = FormWindowState.Normal
		Me.FormBorderStyle = FormBorderStyle.None

		Me.Size = New Point(250, 200)
		Me.BackColor = Color.Gray
		Me.Top = (3 * screenHeight / 4) - (Me.Height / 2)
		Me.Left = (screenWidth / 2) - (Me.Width / 2)

		Me.Controls.Add(Me.contButton)
		xCenter(Me.contButton)

		Me.Controls.Add(Me.subjPanel)
		xCenter(Me.subjPanel, 0.4)

		Me.subjBox.Select()

	End Sub

	Private condN As Integer
	Private subjN As Integer

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If Me.subjBox.Text = "" OrElse Val(Me.subjBox.Text) < 1 OrElse Val(Me.subjBox.Text) > 500 Then 'Interestingly, Val("these are letters") returns 0, and results in error
			MsgBox("Bitte geben Sie eine korrekte VPNr ein!", MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub
		Else
			Me.subjN = CInt(Me.subjBox.Text)
			Me.condN = setCond(Me.subjN)
			Select Case Me.condN
				Case 0
					mainForm.keyAss = "Apos"
					mainForm.firstNames = "Pos"

				Case 1
					mainForm.keyAss = "Apos"
					mainForm.firstNames = "Neg"
				Case 2
					mainForm.keyAss = "Aneg"
					mainForm.firstNames = "Pos"
				Case 3
					mainForm.keyAss = "Aneg"
					mainForm.firstNames = "Neg"
			End Select
			mainForm.dataFrame("Subject") = Me.subjN.ToString
			mainForm.dataFrame("Key") = mainForm.keyAss
			mainForm.dataFrame("FirstNames") = mainForm.firstNames
		End If

		'mainForm.contButton.PerformClick()
		Me.Close()
	End Sub

	Private Sub confirmEnter(sender As Object, e As KeyEventArgs) Handles subjBox.KeyDown
		If e.KeyCode = Keys.Enter Then
			Me.contButton.PerformClick()
		End If
    End Sub

    Private Sub suppressNonNumeric(sender As Object, e As KeyPressEventArgs) Handles subjBox.KeyPress
        If e.KeyChar <> ControlChars.Back AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

End Class