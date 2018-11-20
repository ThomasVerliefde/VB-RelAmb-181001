
Public Class otherForm
	Inherits Form


	Private otherInstr As New instructionBox(horizontalDist:=0.5, verticalDist:=0.25)
	Private WithEvents otherBox1 As New TextBox
	Private WithEvents otherBox2 As New TextBox
	Private otherPanel1 As New labelledBox(Me.otherBox1, "Name der signifikanten anderen 1:", 180)
	Private otherPanel2 As New labelledBox(Me.otherBox2, "Name der signifikanten anderen 2:", 180)
	Private WithEvents contButton As New continueButton
	Private otherPos As New List(Of String)
	Private otherNeg As New List(Of String)

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.Controls.Add(Me.otherInstr)
		xCenter(Me.otherInstr, 0.25)
		Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_other" & mainForm.firstNames)

		Me.Controls.Add(Me.contButton)
		xCenter(Me.contButton, 0.8)

		Me.Controls.Add(Me.otherPanel1)
		xCenter(Me.otherPanel1, 0.5, 0.45)
		Me.Controls.Add(Me.otherPanel2)
		xCenter(Me.otherPanel2, 0.6, 0.45)

		Me.otherBox1.Select()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If Me.otherBox1.Text = "" OrElse Me.otherBox2.Text = "" OrElse
			Not IsName(Me.otherBox1.Text) OrElse Not IsName(Me.otherBox2.Text) OrElse
				Me.otherBox1.Text.Length < 2 OrElse Me.otherBox2.Text.Length < 2 OrElse
				Me.otherBox1.Text.Length > 14 OrElse Me.otherBox2.Text.Length > 14 OrElse
				otherPos.Contains(Me.otherBox1.Text) OrElse otherNeg.Contains(Me.otherBox1.Text) OrElse
				otherPos.Contains(Me.otherBox2.Text) OrElse otherNeg.Contains(Me.otherBox2.Text) OrElse
				Me.otherBox1.Text = Me.otherBox2.Text Then

			MsgBox("Bitte geben Sie gültige Namen ein!" & vbCrLf &
					"[Namen sollten zwischen 2 und 14 Buchstaben lang sein," & vbCrLf &
					"nur Buchstaben enthalten (Umlaute und Bindestriche sind erlaubt), und unverwechselbar sein.]" _
					& vbCrLf, MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub


		ElseIf otherPos.Count = 0 AndAlso otherNeg.Count = 0 Then
			Select Case mainForm.firstNames
				Case "Pos"
					otherPos.AddRange({Me.otherBox1.Text, Me.otherBox2.Text})
					Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherNeg")
				Case "Neg"
					otherNeg.AddRange({Me.otherBox1.Text, Me.otherBox2.Text})
					Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherPos")
			End Select

			Me.otherBox1.Text = ""
			Me.otherBox2.Text = ""
			Me.otherBox1.Select()

		ElseIf otherPos.Count = 2 Then
			mainForm.otherNeg.AddRange({Me.otherBox1.Text, Me.otherBox2.Text})
			mainForm.otherPos.Concat(otherPos)
			Me.Close()
		ElseIf otherNeg.Count = 2 Then
			mainForm.otherPos.AddRange({Me.otherBox1.Text, Me.otherBox2.Text})
			mainForm.otherNeg.Concat(otherNeg)
			Me.Close()
		End If

	End Sub

	Private Sub pressEnter(sender As Object, e As KeyEventArgs) Handles otherBox1.KeyDown, otherBox2.KeyDown
		If e.KeyCode = Keys.Enter Then
			If 0 Then
			ElseIf Me.otherBox1.Text = "" Then
				Me.otherBox1.Select()
			ElseIf Me.otherBox2.Text = "" Then
				Me.otherBox2.Select()
			Else Me.contButton.PerformClick()
			End If
		End If
	End Sub

	Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles otherBox1.KeyPress, otherBox2.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not IsName(Me.otherBox1.Text & e.KeyChar) AndAlso Not IsName(Me.otherBox2.Text & e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class