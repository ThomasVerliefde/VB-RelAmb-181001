
Public Class explicitForm
    Inherits Form


	'Label1.Text = "Wie positiv finden Sie dieses Wort?"
	'Label11.Text = "Konzentrieren Sie sich für Ihr Urteil bitte nur auf die positiven Aspekte des Wortes und ignorieren Sie mögliche negative Aspekte."

	'Label2.Text = "Wie negativ finden Sie dieses Wort?"
	'Label12.Text = "Konzentrieren Sie sich für Ihr Urteil bitte nur auf die negativen Aspekte des Wortes und ignorieren Sie mögliche positive Aspekte."

	'Label3.Text = "Wie hin- und hergerissen fühlen Sie sich angesichts der Bedeutung dieses Wortes?"

	'Label4.Text = "neutral"
	'Label5.Text = "sehr positiv"
	'Label6.Text = "neutral"
	'Label7.Text = "sehr negativ"
	'Label8.Text = "überhaupt nicht"
	'Label9.Text = "sehr"

	Private trackB1 As New TrackBar
	Private labB1 As New labelledTrackbar(Me.trackB1)

	Private trackB2 As New TrackBar
	Private labB2 As New labelledTrackbar(Me.trackB2)

	Private trackB3 As New TrackBar
	Private labB3 As New labelledTrackbar(Me.trackB3)

	Private labIntro As New Label
	Private labName As New Label

	Private WithEvents contButton As New continueButton

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		With Me.labIntro
			.Text = "Please answer these questions about:"
			.Font = sansSerif22
			.Width = TextRenderer.MeasureText(.Text, sansSerif22).Width
			.Height = TextRenderer.MeasureText(.Text, sansSerif22).Height
		End With



		Me.Controls.Add(Me.labIntro)
		xCenter(Me.labIntro, verticalDist:=0.1, setLeft:=50)

		Me.Controls.Add(Me.labName)
		xCenter(Me.labName, verticalDist:=0.1)

		Me.Controls.Add(Me.labB1)
		xCenter(Me.labB1, verticalDist:=0.3)
		Me.Controls.Add(Me.labB2)
		xCenter(Me.labB2, verticalDist:=0.5)
		Me.Controls.Add(Me.labB3)
		xCenter(Me.labB3, verticalDist:=0.7)

		Me.Controls.Add(contButton)
		xCenter(contButton)
		Me.contButton.PerformClick()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		Dim otherName As String = "Frederick"
		Me.contButton.Enabled = False

		'For Each dataFrame()

		With Me.labName
			.Text = otherName
			.Font = sansSerif25B
			.Width = TextRenderer.MeasureText(.Text, sansSerif25B).Width
			.Height = TextRenderer.MeasureText(.Text, sansSerif25B).Height
		End With

		' Positive SRI

		Me.labB1.reInit("How helpful is  " & otherName & "  when you need advice?")
		Me.labB2.reInit("How helpful is  " & otherName & "  when you need understanding?")
		Me.labB3.reInit("How helpful is  " & otherName & "  when you need a favor?")

		' Negative SRI

		Me.labB1.reInit("How upsetting is  " & otherName & "  when you need advice?")
		Me.labB2.reInit("How upsetting is  " & otherName & "  when you need understanding?")
		Me.labB3.reInit("How upsetting is  " & otherName & "  when you need a favor?")

		' Other

		Me.labB1.reInit("Wie positiv finden Sie " & otherName & "?" & vbCrLf & " Konzentrieren Sie sich für Ihr Urteil bitte nur auf die positiven Aspekte" & vbCrLf & "und ignorieren Sie mögliche negative Aspekte.", 0.5, "neutral", "sehr positiv")
		Me.labB2.reInit("Wie negativ finden Sie " & otherName & "?" & vbCrLf & " Konzentrieren Sie sich für Ihr Urteil bitte nur auf die negativen Aspekte" & vbCrLf & "und ignorieren Sie mögliche positive Aspekte.", 0.5, "neutral", "sehr negativ")
		Me.labB3.reInit("Wie hin- und hergerissen fühlen Sie sich angesichts " & otherName,, "überhaupt nicht", "sehr")


	End Sub


End Class