Imports NodaTime

Public Class mainForm
    Inherits Form

    Public dataFrame As New Dictionary(Of String, String)
    Public instructionCount As New Integer

    Friend WithEvents contButton As New continueButton
    Private ReadOnly instrText As New instructionBox

    Private time As IClock = SystemClock.Instance

    Public timeStart As Instant = time.GetCurrentInstant
    Public timeOther As Duration
    Public timePractice As Duration
    Public timeExperiment As Duration
    Public timeExplicit As Duration
    Public timeDemographics As Duration
    Public timeTotal As Duration

    Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = FormWindowState.Maximized
        FormBorderStyle = FormBorderStyle.None
        BackColor = Color.White

        If Not dataFrame.ContainsKey("Subject") Then
            subjectForm.ShowDialog()
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
                instrText.Text = My.Resources.ResourceManager.GetString("_2_practice" & subjectForm.keyAss)
                'practiceForm.ShowDialog()

            Case 3 'Experiment Proper
                instrText.Text = My.Resources.ResourceManager.GetString("_3_experiment" & subjectForm.keyAss)
                'expForm.ShowDialog()

            Case 4 'Explicit Measurements of Ambivalence
                instrText.Text = My.Resources.ResourceManager.GetString("_4_explicitInstr")
                'explicitForm.ShowDialog()

            Case 5 'Demographic Information
                instrText.Text = My.Resources.ResourceManager.GetString("_5_demoInstr")
                'demoForm.ShowDialog()

            Case 6 'End of Experiment

                instrText.Text = My.Resources.ResourceManager.GetString("_6_endInstr")
                instrText.Font = New Font("Microsoft Sans Serif", 40)
                instrText.TextAlign = HorizontalAlignment.Center
                contButton.Text = "Abbrechen"

                timeTotal = time.GetCurrentInstant - timeStart





                dataFrame("timeOther") = timeOther.TotalMinutes.ToString
                dataFrame("timePractice") = timePractice.TotalMinutes.ToString
                dataFrame("timeExperiment") = timeExperiment.TotalMinutes.ToString
                dataFrame("timeExplicit") = timeExplicit.TotalMinutes.ToString
                dataFrame("timeDemographics") = timeDemographics.TotalMinutes.ToString
                dataFrame("timeTotal") = timeTotal.TotalMinutes.ToString
                dataFrame("hostName") = Net.Dns.GetHostName()

                '' Copy of "Ende.vb"

                '' Die Verweildauer auf den einzelnen Seiten und Overall abspeichern
                'dataFrame("timeAPDeliver") = timeAPDeliver.ToString
                'dataFrame("timeAPBerger") = timeAPBerger.ToString
                'dataFrame("timeExplicit") = timeExplicit.ToString
                'dataFrame("timePRDeliver") = timePRDeliver.ToString
                'dataFrame("timePRBerger") = timePRBerger.ToString
                'dataFrame("timeOverall") = difference.ToString

                '' Stimulus-Listen der Primingphasen im Datenfile ablegen
                'dataFrame("RF_practiceDL_primes") = rf1
                'dataFrame("RF_practiceDL_targets") = rf2
                'dataFrame("RF_practiceBH_primes") = rf3
                'dataFrame("RF_practiceBH_targets") = rf4
                'dataFrame("RF_targetsBH") = rf8
                'dataFrame("RF_primesDL") = rf5
                'dataFrame("RF_targetsDL") = rf6
                'dataFrame("RF_primesBH") = rf7

                saveCSV(dataFrame) 'Optionally, you can specify a specific path + filename as a second argument, otherwise it will automatically save it as rawData.csv in the main folder

            Case Else
                Close()
        End Select
        instructionCount += 1
    End Sub





End Class
