
Public Class practiceForm
	Inherits Form

	Private stopwatchTarget As New Stopwatch
	Private answeringTime As Long

	Private WithEvents timerITI As New Timer
	Private WithEvents timerFix As New Timer
	Private WithEvents timerPrime As New Timer

	Private leftLab As New Label
	Private rightLab As New Label
	Private slowLab As New Label
	Private fixLab As New Label
	Private primeLab As New Label
	Private targetLab As New Label

	Private ignoreKeys As Boolean
	Private trialCounter As Integer

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.timerITI.Interval = 1500
		Me.timerFix.Interval = 500
		Me.timerPrime.Interval = 150

		Select Case mainForm.keyAss
			Case "Apos"
				Me.leftLab.Text = "A = positiv"
				Me.rightLab.Text = "L = negativ"
			Case "Aneg"
				Me.leftLab.Text = "A = negativ"
				Me.rightLab.Text = "L = positiv"
		End Select

		With Me.leftLab
			.Left = 100
			.Top = 100
			.Font = sansSerif20
			.TextAlign = ContentAlignment.TopLeft
			.AutoSize = True
		End With

		With Me.rightLab
			'.Anchor = AnchorStyles.Top Or AnchorStyles.Right
			.Left = Me.Width - TextRenderer.MeasureText(Me.rightLab.Text, sansSerif20).Width - 100
			.Top = 100
			.Font = sansSerif20
			.TextAlign = ContentAlignment.TopRight
			.AutoSize = True
		End With

		With Me.slowLab
			.Visible = False
			.Text = "Schneller!"
			.AutoSize = True
			.Font = sansSerif25B
			.ForeColor = Color.Red
			.TextAlign = ContentAlignment.TopCenter
			.AutoSize = True
		End With
		xCenter(Me.slowLab, 0.5)

		With Me.fixLab
			.Visible = False
			.Text = "+"
			.Font = sansSerif72
			.TextAlign = ContentAlignment.TopCenter
			.AutoSize = True
		End With
		xCenter(Me.fixLab, 0.5)

		With Me.primeLab
			.Visible = False
			.Text = ""
			.Font = sansSerif60
			.TextAlign = ContentAlignment.TopCenter
			.AutoSize = True
		End With
		xCenter(Me.primeLab, 0.5)

		With Me.targetLab
			.Visible = False
			.Text = ""
			.Font = sansSerif60
			.TextAlign = ContentAlignment.TopCenter
			.AutoSize = True
		End With
		xCenter(Me.targetLab, 0.5)

		Me.Controls.AddRange({Me.leftLab, Me.rightLab, Me.slowLab, Me.fixLab, Me.primeLab, Me.targetLab})

		Me.timerITI.Enabled = True

	End Sub

	Private Sub timerITI_Tick(sender As Object, e As EventArgs) Handles timerITI.Tick
		Me.timerITI.Stop()
		Me.timerFix.Start()
		Me.slowLab.Visible = False
		Me.fixLab.Visible = True
	End Sub

	Private Sub timerFix_Tick(sender As Object, e As EventArgs) Handles timerFix.Tick
		Me.timerFix.Stop()
		Me.timerPrime.Start()
		Me.primeLab.Text = practiceTrials(trialCounter)(0)
		Me.fixLab.Visible = False
		Me.fixLab.Visible = False
		Me.primeLab.Visible = True
	End Sub

	Private Sub timerPrime_Tick(sender As Object, e As EventArgs) Handles timerPrime.Tick
		Me.timerPrime.Stop()
		Me.stopwatchTarget.Start()
		Me.targetLab.Text = practiceTrials(trialCounter)(0)
		Me.primeLab.Visible = False
		Me.targetLab.Visible = True
		Me.ignoreKeys = False
	End Sub

	Private Sub responseAL(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

		e.Handled = Me.ignoreKeys 'Stops the event when it is not the categorisation-phase

		If e.KeyCode = Keys.A Or e.KeyCode = Keys.L Then 'Non-short-circuited "Or" (instead of OrElse) to prevent systematic differences between A and L (although these would be minimal)
			Me.answeringTime = Me.stopwatchTarget.ElapsedMilliseconds
			Me.stopwatchTarget.Reset()
			Me.timerITI.Start()
			Me.targetLab.Visible = False
			Me.slowLab.Visible = Me.answeringTime > 1500
			Me.ignoreKeys = True

			dataFrame("practice_" & trialCounter & "_answer") = e.KeyCode.ToString
			dataFrame("practice_" & trialCounter & "_time") = answeringTime.ToString
			dataFrame("practice_" & trialCounter & "_prime") = practiceTrials(trialCounter)(0)
			dataFrame("practice_" & trialCounter & "_target") = practiceTrials(trialCounter)(1)

			Me.trialCounter += 1

			If Me.trialCounter = practiceTrials.Count Then
				Me.Close()
			End If
		End If

	End Sub

End Class