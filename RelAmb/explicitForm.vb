
Public Class explicitForm
	Inherits Form

	Private WithEvents trackB1 As New TrackBar
	Private labB1 As New labelledTrackbar(Me.trackB1)

	Private WithEvents trackB2 As New TrackBar
	Private labB2 As New labelledTrackbar(Me.trackB2)

	Private WithEvents trackB3 As New TrackBar
	Private labB3 As New labelledTrackbar(Me.trackB3)

	Private labIntro As New Label
	Private labName As New Label

	Private WithEvents numText As New TextBox
	Private numBox As New labelledBox(Me.numText, "Wie viele Leute kennen Sie mit diesem Vornamen?")

	Private WithEvents contButton As New continueButton
	Private otherKeys As New List(Of String)({"otherPos1", "otherPos2", "otherNeg1", "otherNeg2"})
	Private otherCount As Integer
	Private questionCount As Integer

	Private otherKey As String
	Private otherName As String

	Private tempFrame As New SortedDictionary(Of String, String)

	Private initB1 As Boolean
	Private B1 As Boolean
	Private B2 As Boolean
	Private B3 As Boolean

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.trackB1.Name = "B1"
		Me.trackB2.Name = "B2"
		Me.trackB3.Name = "B3"

		shuffleList(Me.otherKeys)

		With Me.labIntro
			.Text = "Bitte beantworten Sie die folgenden Fragen zu:"
			.Font = sansSerif22
			.Width = TextRenderer.MeasureText(.Text, sansSerif22).Width
			.Height = TextRenderer.MeasureText(.Text, sansSerif22).Height
		End With

		Me.Controls.Add(Me.labIntro)
		xCenter(Me.labIntro, verticalDist:=0.1, setLeft:=50)

		Me.Controls.Add(Me.labName)
		xCenter(Me.labName, verticalDist:=0.095)

		Me.Controls.Add(Me.labB1)
		Me.Controls.Add(Me.labB2)
		Me.Controls.Add(Me.labB3)
		Me.Controls.Add(Me.numBox)
		xCenter(Me.labB1, verticalDist:=0.3)
		xCenter(Me.labB2, verticalDist:=0.5)
		xCenter(Me.labB3, verticalDist:=0.7)
		xCenter(Me.numBox, 0.45, 0.4)

		Me.Controls.Add(Me.contButton)
		xCenter(Me.contButton)

		Me.contButton.PerformClick()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click



		If Me.otherCount >= Me.otherKeys.Count Then

			For Each item In Me.tempFrame
				dataFrame.Add(item.Key, item.Value)
			Next

			Me.Close()

		Else

			If Me.questionCount <> 0 Then 'Saving all the data (except the initial time); utilises the previous iterations of otherKey
				Select Case (Me.questionCount - 1) Mod 4 'Because this triggers at the start of a new buttonpush, events are lagged by 1

					Case 0 ' Positive SRI

						Me.tempFrame(Me.otherKey & "_SRI_Pos1_Adv") = Me.trackB1.Value.ToString
						Me.tempFrame(Me.otherKey & "_SRI_Pos2_Und") = Me.trackB2.Value.ToString
						Me.tempFrame(Me.otherKey & "_SRI_Pos3_Fav") = Me.trackB3.Value.ToString

					Case 1 ' Negative SRI		

						Me.tempFrame(Me.otherKey & "_SRI_Neg1_Adv") = Me.trackB1.Value.ToString
						Me.tempFrame(Me.otherKey & "_SRI_Neg2_Und") = Me.trackB2.Value.ToString
						Me.tempFrame(Me.otherKey & "_SRI_Neg3_Fav") = Me.trackB3.Value.ToString

					Case 2
						' Other

						Me.tempFrame(Me.otherKey & "_Dir_Pos") = Me.trackB1.Value.ToString
						Me.tempFrame(Me.otherKey & "_Dir_Neg") = Me.trackB2.Value.ToString
						Me.tempFrame(Me.otherKey & "_Dir_Amb") = Me.trackB3.Value.ToString

					Case 3
						' How many?
						Me.tempFrame(Me.otherKey & "_Num") = Me.numText.Text

				End Select

			End If

			Me.otherKey = Me.otherKeys(Me.otherCount)
			Me.otherName = dataFrame(Me.otherKey)

			With Me.labName
				.Text = Me.otherName
				.Font = sansSerif25B
				.Width = TextRenderer.MeasureText(.Text, sansSerif25B).Width
				.Height = TextRenderer.MeasureText(.Text, sansSerif25B).Height
			End With

			'Resetting the checks whether the trackbars have received attention
			Me.initB1 = False
			Me.B1 = False
			Me.B2 = False
			Me.B3 = False
			Me.contButton.Enabled = False

			Select Case Me.questionCount Mod 4

				Case 0

					Me.labB1.Visible = True
					Me.labB2.Visible = True
					Me.labB3.Visible = True
					Me.numBox.Visible = False

					' Positive SRI
					' How helpful is XXX when you need advice/understanding/a favour? | willing to help, or useful

					Me.labB1.reInit("Wie hilfreich ist " & Me.otherName & ", wenn Sie Rat brauchen?", Me.trackB1)
					Me.labB2.reInit("Wie hilfreich ist " & Me.otherName & ", wenn Sie Verständnis brauchen?", Me.trackB2)
					Me.labB3.reInit("Wie hilfreich ist " & Me.otherName & ", wenn Sie einen Gefallen brauchen?", Me.trackB3)

				Case 1

					' Negative SRI
					' How upsetting is XXX when you need advice/understanding/a favour | upsetting: making someone feel worried, unhappy, or angry (cambridge dictionary)
					' Should come up with a better word still
					' Kathis suggestion: "in Aufregung versetzen"
					' Max' suggestion: "erschütternd"
					' Other interesting ideas: "ärgerlich", "umständlich", "beunruhigend", "erschwerend", "vesrchlimmernd", "agitierend", "verwirrend"

					Me.labB1.reInit("Wie umständlich ist " & Me.otherName & ", wenn Sie Rat brauchen?", Me.trackB1)
					Me.labB2.reInit("Wie umständlich ist " & Me.otherName & ", wenn Sie Verständnis brauchen?", Me.trackB2)
					Me.labB3.reInit("Wie umständlich ist " & Me.otherName & ", wenn Sie einen Gefallen brauchen?", Me.trackB3)

				Case 2 ' Explicit positive, negative, and ambivalent

					Me.labB1.reInit("Wie positiv finden Sie " & Me.otherName & "?" & vbCrLf &
									" Konzentrieren Sie sich für Ihr Urteil bitte nur auf die positiven Aspekte" & vbCrLf &
									"und ignorieren Sie mögliche negative Aspekte.", Me.trackB1, 0.5, "neutral", "sehr positiv")
					Me.labB2.reInit("Wie negativ finden Sie " & Me.otherName & "?" & vbCrLf &
									" Konzentrieren Sie sich für Ihr Urteil bitte nur auf die negativen Aspekte" & vbCrLf &
									"und ignorieren Sie mögliche positive Aspekte.", Me.trackB2, 0.5, "neutral", "sehr negativ")
					Me.labB3.reInit("Wie hin- und hergerissen fühlen Sie sich angesichts " & Me.otherName, Me.trackB3,, "überhaupt nicht", "sehr")

				Case 3 ' How many do you know?

					Me.labB1.Visible = False
					Me.labB2.Visible = False
					Me.labB3.Visible = False
					Me.numBox.Visible = True

					Me.numText.Text = ""
					Me.numText.Focus()

					Me.otherCount += 1

			End Select

			Me.questionCount += 1

		End If

	End Sub

	Private Sub enableTrack(sender As Object, e As EventArgs) Handles trackB1.MouseDown, trackB2.MouseDown, trackB3.MouseDown

		'Console.WriteLine(sender)

		'Console.WriteLine(DirectCast(sender, TrackBar).Name)

		Select Case DirectCast(sender, TrackBar).Name
			Case "B1"
				Me.B1 = True

			Case "B2"
				Me.B2 = True

			Case "B3"
				Me.B3 = True

		End Select

		If Me.B1 AndAlso Me.B2 AndAlso Me.B3 Then
			Me.contButton.Enabled = True
		End If

	End Sub

	Private Sub enableNum(sender As Object, e As EventArgs) Handles numText.TextChanged
		If Val(Me.numText.Text) < 1 OrElse Val(Me.numText.Text) > 25 Then
			Me.contButton.Enabled = False
		ElseIf Val(Me.numText.Text) > 0 Then
			Me.contButton.Enabled = True
		End If
	End Sub

	Private Sub confirmEnter(sender As Object, e As KeyEventArgs) Handles numText.KeyDown
		If e.KeyCode = Keys.Enter Then
			Me.contButton.PerformClick()
		End If
	End Sub

	Private Sub suppressNonNumeric(sender As Object, e As KeyPressEventArgs) Handles numText.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not IsNumeric(e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class