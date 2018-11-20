
Public Class demographicsForm
    Inherits Form

	Private listBox1 As New ListBox

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.Controls.Add(listBox1)
		xCenter(listBox1, 0.25)

	End Sub


	'mainForm.contButton.PerformClick()


End Class