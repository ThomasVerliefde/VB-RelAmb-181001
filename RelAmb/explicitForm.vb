
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

	Private barPos1 As New TrackBar
	Private posSRI1 As New labelledTrackbar(Me.barPos1, "", "überhaupt nicht", "sehr")

	Private barPos2 As New TrackBar
	Private posSRI2 As New labelledTrackbar(Me.barPos2, "", "überhaupt nicht", "sehr")

	Private barPos3 As New TrackBar
	Private posSRI3 As New labelledTrackbar(Me.barPos3, "", "überhaupt nicht", "sehr")

	Private barNeg1 As New TrackBar
	Private negSRI1 As New labelledTrackbar(Me.barNeg1, "", "überhaupt nicht", "sehr")

	Private barNeg2 As New TrackBar
	Private negSRI2 As New labelledTrackbar(Me.barNeg2, "", "überhaupt nicht", "sehr")

	Private barNeg3 As New TrackBar
	Private negSRI3 As New labelledTrackbar(Me.barNeg3, "", "überhaupt nicht", "sehr")

	Private WithEvents contButton As New continueButton

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White


		Me.Controls.Add(Me.posSRI1)
		xCenter(Me.posSRI1, verticalDist:=0.4)

		Me.Controls.Add(contButton)
		xCenter(contButton)
		Me.contButton.PerformClick()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		Me.contButton.Enabled = False

		Me.posSRI1.reLabel("How helpful is  " & "NAME" & "  when you need advice?")
		Me.posSRI2.reLabel("How helpful is  " & "NAME" & "  when you need understanding?")
		Me.posSRI3.reLabel("How helpful is  " & "NAME" & "  when you need a favor?")


		Me.negSRI1.reLabel("How upsetting is  " & "NAME" & "  when you need advice?")
		Me.negSRI2.reLabel("How upsetting is  " & "NAME" & "  when you need understanding?")
		Me.negSRI3.reLabel("How upsetting is  " & "NAME" & "  when you need a favor?")


	End Sub


End Class