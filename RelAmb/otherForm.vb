
Public Class otherForm
	Inherits Form


	Private otherInstr As New instructionBox(horizontalDist:=0.5, verticalDist:=0.25)
	Private WithEvents otherBox1 As New TextBox
	Private WithEvents otherBox2 As New TextBox
	Private otherPanel1 As New labelledBox(otherBox1, "Name der signifikanten anderen 1:", 180)
	Private otherPanel2 As New labelledBox(otherBox2, "Name der signifikanten anderen 2:", 180)
	Private WithEvents contButton As New continueButton
	'Friend mainForm.otherPos As New List(Of String)
	'Friend mainForm.otherNeg As New List(Of String)

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		WindowState = FormWindowState.Maximized
		FormBorderStyle = FormBorderStyle.None
		BackColor = Color.White

		Controls.Add(otherInstr)
		xCenter(otherInstr, 0.25)
		otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_other" & mainForm.firstNames)
		'otherInstr.TextAlign = HorizontalAlignment.Center

		Controls.Add(contButton)
		xCenter(contButton, 0.8)

		Controls.Add(otherPanel1)
		xCenter(otherPanel1, 0.5, 0.45)
		Controls.Add(otherPanel2)
		xCenter(otherPanel2, 0.6, 0.45)

		otherBox1.Select()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If otherBox1.Text = "" OrElse otherBox2.Text = "" OrElse
			Not IsName(otherBox1.Text) OrElse Not IsName(otherBox2.Text) OrElse
				otherBox1.Text.Length < 2 OrElse otherBox2.Text.Length < 2 OrElse
				otherBox1.Text.Length > 14 OrElse otherBox2.Text.Length > 14 OrElse
				mainForm.otherPos.Contains(otherBox1.Text) OrElse mainForm.otherNeg.Contains(otherBox1.Text) OrElse
				mainForm.otherPos.Contains(otherBox2.Text) OrElse mainForm.otherNeg.Contains(otherBox2.Text) OrElse
				otherBox1.Text = otherBox2.Text Then

			MsgBox("Bitte geben Sie gültige Namen ein!" & vbCrLf &
					"[Namen sollten zwischen 2 und 14 Buchstaben lang sein," & vbCrLf &
					"nur Buchstaben enthalten (Umlaute und Bindestriche sind erlaubt), und unverwechselbar sein.]" _
					& vbCrLf, MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub


		ElseIf mainForm.otherPos.Count = 0 AndAlso mainForm.otherNeg.Count = 0 Then
			Select Case mainForm.firstNames
				Case "Pos"
					mainForm.otherPos.AddRange({otherBox1.Text, otherBox2.Text})
					otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherNeg")
				Case "Neg"
					mainForm.otherNeg.AddRange({otherBox1.Text, otherBox2.Text})
					otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherPos")
			End Select

			otherBox1.Text = ""
			otherBox2.Text = ""
			otherBox1.Select()

		ElseIf mainForm.otherPos.Count = 2 Then
			mainForm.otherNeg.AddRange({otherBox1.Text, otherBox2.Text})
			Close()
		ElseIf mainForm.otherNeg.Count = 2 Then
			mainForm.otherPos.AddRange({otherBox1.Text, otherBox2.Text})
			Close()
		End If

	End Sub

	Private Sub pressEnter(sender As Object, e As KeyEventArgs) Handles otherBox1.KeyDown, otherBox2.KeyDown
		If e.KeyCode = Keys.Enter Then
			If 0 Then
			ElseIf otherBox1.Text = "" Then
				otherBox1.Select()
			ElseIf otherBox2.Text = "" Then
				otherBox2.Select()
			Else contButton.PerformClick()
			End If
		End If
	End Sub

	Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles otherBox1.KeyPress, otherBox2.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not IsName(otherBox1.Text & e.KeyChar) AndAlso Not IsName(otherBox2.Text & e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class