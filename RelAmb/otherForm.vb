
Public Class otherForm
    Inherits Form


    Private otherInstr As New instructionBox
    Private WithEvents otherBox1 As New TextBox
    Private WithEvents otherBox2 As New TextBox
    Private otherPanel1 As New labelledBox(otherBox1, "S.A. 1", 400)
    Private otherPanel2 As New labelledBox(otherBox1, "S.A. 2", 400)
    Private WithEvents contButton As continueButton

    Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = FormWindowState.Maximized
        FormBorderStyle = FormBorderStyle.None
        BackColor = Color.White

        Controls.Add(otherInstr)
        xCenter(otherInstr)
        otherInstr.Text = My.Resources.ResourceManager.GetString("_1_otherPos")

        Controls.Add(contButton)
        xCenter(contButton)

        Controls.Add(otherPanel1)
        xCenter(otherPanel1, 0.4)
        Controls.Add(otherPanel2)
        xCenter(otherPanel2, 0.6)

        otherBox1.Select()

    End Sub

    Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

        If otherBox1.Text = "" OrElse otherBox2.Text = "" OrElse Not IsAlpha(otherBox1.Text) _
            OrElse Not IsAlpha(otherBox2.Text) OrElse otherBox1.Text.Length = 1 OrElse otherBox2.Text.Length = 1 Then
            MsgBox("Bitte geben Sie gültige Namen ein!", MsgBoxStyle.Critical, Title:="Fehler!")
            Exit Sub
        Else

        End If

        mainForm.contButton.PerformClick()
        Close()
    End Sub

    Private Sub confirmEnter(sender As Object, e As KeyEventArgs) Handles otherBox1.KeyDown, otherBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            contButton.PerformClick()
        End If
    End Sub

    Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles otherBox1.KeyPress, otherBox2.KeyPress
        If e.KeyChar <> ControlChars.Back AndAlso Not IsAlpha(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub


End Class