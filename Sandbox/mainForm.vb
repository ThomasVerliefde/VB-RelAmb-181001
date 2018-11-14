Imports NodaTime

Public Class mainForm
    Inherits Form

    Public dataFrame As New Dictionary(Of String, String)
    Public instructionCount As New Integer

    Friend WithEvents contButton As New continueButton
    Private ReadOnly instrText As New instructionBox

    Private time As IClock = SystemClock.Instance
    Private startT As Instant = time.GetCurrentInstant
    Private endT As Instant
    Private totalT As Duration

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
            Case 0 'Start of Experiment
                instrText.Text = My.Resources.ResourceManager.GetString("_0_mainInstr")

            Case 1 'Collecting Names of 'Significant Others'
                instrText.Text = My.Resources.ResourceManager.GetString("_1_otherInstr")
                'otherForm.ShowDialog()

            Case 2 'Practice Trials
                instrText.Text = My.Resources.ResourceManager.GetString("_2_practice" & subjForm.keyAss)
                'practiceForm.ShowDialog()

            Case 3 'Experiment Proper
                instrText.Text = My.Resources.ResourceManager.GetString("_3_experiment" & subjForm.keyAss)
                'expForm.ShowDialog()

            Case 4 'Explicit Measurements of Ambivalence
                instrText.Text = My.Resources.ResourceManager.GetString("_4_explicitInstr")
                'explicitForm.ShowDialog()

            Case 5 'Demographic Information
                instrText.Text = My.Resources.ResourceManager.GetString("_5_demoInstr")
                'demoForm.ShowDialog()

            Case 6 'End of Experiment


                endT = time.GetCurrentInstant
                totalT = endT - startT

                instrText.Text = totalT.TotalMinutes.ToString

                'instrText.Text = My.Resources.ResourceManager.GetString("_6_endInstr")
                instrText.Font = New Font("Microsoft Sans Serif", 40)
                instrText.TextAlign = HorizontalAlignment.Center
                contButton.Text = "Abbrechen"

                '' Copy of "Ende.vb"

                'Dim strHostName As String
                '' PC Name
                'strHostName = System.Net.Dns.GetHostName()

                '' Total Time

                'Dim endT As New Instant
                'Dim totT As New Duration

                'totT = startT - endT
                'totT.TotalMinutes


                'VL.time2 = Format(Now, "Long Time")
                'VL.difference = VL.time2 - VL.time1

                '' Die Verweildauer auf den einzelnen Seiten und Overall abspeichern
                'VL.data("timeAPDeliver") = VL.timeAPDeliver.ToString
                'VL.data("timeAPBerger") = VL.timeAPBerger.ToString
                'VL.data("timeExplicit") = Explicit_Berger.timeExplicit.ToString
                'VL.data("timePRDeliver") = VL.timePRDeliver.ToString
                'VL.data("timePRBerger") = VL.timePRBerger.ToString
                'VL.data("timeOverall") = VL.difference.ToString

                '' Stimulus-Listen der Primingphasen im Datenfile ablegen
                'VL.data("RF_practiceDL_primes") = VL.rf1
                'VL.data("RF_practiceDL_targets") = VL.rf2
                'VL.data("RF_practiceBH_primes") = VL.rf3
                'VL.data("RF_practiceBH_targets") = VL.rf4
                'VL.data("RF_primesDL") = VL.rf5
                'VL.data("RF_targetsDL") = VL.rf6
                'VL.data("RF_primesBH") = VL.rf7
                'VL.data("RF_targetsBH") = VL.rf8

                saveCSV(dataFrame)

            Case Else
                Close()
        End Select
        instructionCount += 1
    End Sub





End Class
