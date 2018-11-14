Public Class mainForm
    Inherits Form

    Public dataFrame As New Dictionary(Of String, String)
    Public instructionCount As New Integer
    Friend WithEvents contButton As New continueButton
    Private ReadOnly instrText As New instructionBox

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = FormWindowState.Maximized
        FormBorderStyle = FormBorderStyle.None
        BackColor = Color.White

        If Not dataFrame.ContainsKey("Subject") Then
            subjForm.ShowDialog()
        End If

        Controls.Add(instrText)
        xCenter(instrText, 0.4)

        Controls.Add(contButton)
        xCenter(contButton)

    End Sub

    Public Sub loadNext(sender As Object, e As EventArgs) Handles contButton.Click
        Select Case instructionCount
            Case 0
                instrText.Text = My.Resources.ResourceManager.GetString("_0_mainInstr")

            Case 1
                instrText.Text = My.Resources.ResourceManager.GetString("_1_otherInstr")

            Case 2
                instrText.Text = My.Resources.ResourceManager.GetString("_2_practice" & subjForm.keyAss)

            Case 3
                instrText.Text = My.Resources.ResourceManager.GetString("_3_experiment" & subjForm.keyAss)

            Case 4
                instrText.Text = My.Resources.ResourceManager.GetString("_4_explicitInstr")

            Case 5
                instrText.Text = My.Resources.ResourceManager.GetString("_5_demoInstr")

            Case 6
                instrText.Text = My.Resources.ResourceManager.GetString("_6_endInstr")
                instrText.Font = New Font("Microsoft Sans Serif", 40)
                instrText.TextAlign = HorizontalAlignment.Center
                contButton.Text = "Abbrechen"
            Case Else
                'saveCSV(dataFrame)
                Close()
        End Select
        instructionCount += 1
    End Sub





End Class
