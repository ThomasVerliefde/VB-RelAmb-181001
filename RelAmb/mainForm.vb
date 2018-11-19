Imports NodaTime

Public Class mainForm
    Inherits Form

    Public dataFrame As New Dictionary(Of String, String) 'main dataframe to save all our data, gets written out at the end of the experiment
    Public instructionCount As New Integer 'main counter, to control the flow of the experiment; increases after each continuebutton press in this form

    Friend WithEvents contButton As New continueButton 'main button, to advance the flow of the experiment
    Private ReadOnly instrText As New instructionBox 'main method of displaying the instructions to participants; disabled & readonly

    Friend time As IClock = SystemClock.Instance 'NodaTime clock instance, which keeps time, at the start of every part, gets saved in a variable (see under)

    'All NodaTime.Instant variables, to check starting points of each part
    Private startT As Instant
    Private otherT As Instant
    Private practiceT As Instant
    Private experimentT As Instant
    Private explicitT As Instant
    Private demographicsT As Instant
    Private endT As Instant

    'All NodaTime.Duration variables, to check the actual duration (difference in sequential starting points) of each part
    Private timeOther As Duration
    Private timePractice As Duration
    Private timeExperiment As Duration
    Private timeExplicit As Duration
    Private timeDemographics As Duration
    Private timeTotal As Duration

    'Variable necessary for grabbing the correct instruction sheet, depending on whether the 'A' key is used to categorize positive adjectives, or for negative adjectives
    Friend keyAss As String
    Friend firstNames As String

    Friend otherPos As New List(Of String)
    Friend otherNeg As New List(Of String)



    '####################################'
    'Imported CODE - Not Yet Consolidated'

    'Variablen für die Übungsphasen

    Public practiceBH = New String(7, 3) {} ' Berger & Hütter Paradigma

    ' Aufgespaltete RF für Datenspeicherung

    Public practiceBHprimes = New String(7) {}
    Public practiceBHtargets = New String(7) {}

    Public practiceposadj = New String() {"geduldig", "zärtlich", "humorvoll", "fleissig"}
    Public practicenegadj = New String() {"boshaft", "korrupt", "ungerecht", "gehässig"}

    Public practiceposnoun = New String() {"Kuss", "Idee", "Chance"}
    Public practicenegnoun = New String() {"Mord", "Angst", "Sklave"}
    Public practiceambnoun = New String() {"Arbeit", "Macht", "Diät"}


    ' Variablen für affektives Priming deLiver/ Berger & Hütter

    ' Wortmaterial Nomen
    Public nounspos = New String() {"Mut", "Lust", "Glück", "Freude"}
    Public nounsneg = New String() {"Leid", "Ärger", "Furcht", "Trauer"}
    Public nounsnonwords = New String() {"hdreeh", "izmcsg", "tdkesu", "unerit", "ilbuel", "rdnslb", "ursdcu", "nsedni"}
    Public nounsamb = New String() {"Karriere", "Regulierung", "Arztbesuch", "Protest", "Alkohol", "Feuer", "Smartphone", "Stolz"}

    ' Wortmaterial Adjektive
    Public adjpos = New String() {"frei", "klug", "treu", "gesund", "beliebt", "ehrlich", "herzlich", "friedlich"}
    Public adjneg = New String() {"blöd", "dumm", "fies", "brutal", "grausam", "neidisch", "peinlich", "widerlich"}
    Public adjnonwords = New String() {"takatg", "ineifk", "aahrte", "ndwrow", "rurtum", "worenl", "iasweg", "enoers"}

    ' Liste mit allen Nomen für die expliziten Ratings
    ' Keine Nonwords!
    Public nouns = New String(15, 1) {}




    ' Material Berger & Hütter (96 Trials)
    ' PRIMES = NOUNS: amb (8), neutral (8), pos (4), neg (4)
    ' TARGETS = ADJECTIVES: pos (8), neg (8)
    ' 24 Primes * 2 Target-Types = 48 Trials * 2 (höhere Reliabilität)
    Public trialsBH = New String(95, 3) {}
    Public primesBH = New String(95) {}
    Public targetsBH = New String(95) {}

    Public rf1 As String ' Practice trials primes
    Public rf2 As String ' Practice trials targets
    Public rf3 As String ' Test trials primes
    Public rf4 As String ' Test trials targets
    Public rf5 As String ' ...
    Public rf6 As String
    Public rf7 As String
    Public rf8 As String

    '########################################################'


    Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = FormWindowState.Maximized
        FormBorderStyle = FormBorderStyle.None
        BackColor = Color.White

        'If Not dataFrame.ContainsKey("Subject") Then
        subjectForm.ShowDialog()
        'End If

        Controls.Add(instrText)
        xCenter(instrText, 0.4)
        instrText.Text = My.Resources.ResourceManager.GetString("_0_mainInstr")

        Controls.Add(contButton)
        xCenter(contButton)

    End Sub

    Public Sub loadNext(sender As Object, e As EventArgs) Handles contButton.Click
        Select Case instructionCount
            Case 0 'Start of Experiment
                subjectForm.Dispose()
                startT = time.GetCurrentInstant()
                instrText.Text = My.Resources.ResourceManager.GetString("_1_otherInstr")

                ''Debug -> Check correct functioning "createPrimes" function
                'Dim otherPrimes = New List(Of String)({"Jan", "Jef", "Dries", "Droes"})
                ''Dim posList = New List(Of String)({"3P", "4P", "5P", "6P"})
                ''Dim negList = New List(Of String)({"3N", "4N", "5N", "6N"})
                ''Dim strList = New List(Of String)({"AAA", "BBB", "CCC", "DDD", "AAAA", "BBBB", "CCCC", "DDDD", "AAAAA", "BBBBB", "CCCCC", "DDDDD", "AAAAAA", "BBBBBB", "CCCCCC", "DDDDDD"})
                'Dim posList = New List(Of String)(My.Resources.experimentPrime_Pos.Split(" "))
                'Dim negList = New List(Of String)(My.Resources.experimentPrime_Neg.Split(" "))
                'Dim strList = New List(Of String)(My.Resources.experimentPrime_Str.Split(" "))
                'Dim test = createPrimes(otherPrimes, posList, negList, strList)
                'For Each primegroup In test
                '    For Each item In primegroup
                '        Console.Write(item + " ")
                '    Next
                'Next
                'End Debug

            Case 1 'Collecting Names of 'Significant Others'
                otherT = time.GetCurrentInstant()
                otherForm.ShowDialog()

                instrText.Text = My.Resources.ResourceManager.GetString("_2_practice" & keyAss)
            Case 2 'Practice Trials
                practiceT = time.GetCurrentInstant()
                'practiceForm.ShowDialog()

                instrText.Text = My.Resources.ResourceManager.GetString("_3_experiment" & keyAss)
            Case 3 'Experiment Proper
                experimentT = time.GetCurrentInstant()
                'expForm.ShowDialog()

                instrText.Text = My.Resources.ResourceManager.GetString("_4_explicitInstr")
            Case 4 'Explicit Measurements of Ambivalence
                explicitT = time.GetCurrentInstant()
                'explicitForm.ShowDialog()

                instrText.Text = My.Resources.ResourceManager.GetString("_5_demoInstr")
            Case 5 'Demographic Information
                demographicsT = time.GetCurrentInstant()
                'demoForm.ShowDialog()

            Case 6 'End of Experiment
                instrText.Text = My.Resources.ResourceManager.GetString("_6_endInstr")
                instrText.Font = New Font("Microsoft Sans Serif", 40)
                'instrText.TextAlign = HorizontalAlignment.Center
                contButton.Text = "Abbrechen"
                endT = time.GetCurrentInstant()

                timeOther = practiceT - otherT
                timePractice = experimentT - practiceT
                timeExperiment = explicitT - experimentT
                timeExplicit = demographicsT - explicitT
                timeDemographics = endT - demographicsT
                timeTotal = endT - startT

                dataFrame("timeOther") = timeOther.TotalMinutes.ToString
                dataFrame("timePractice") = timePractice.TotalMinutes.ToString
                dataFrame("timeExperiment") = timeExperiment.TotalMinutes.ToString
                dataFrame("timeExplicit") = timeExplicit.TotalMinutes.ToString
                dataFrame("timeDemographics") = timeDemographics.TotalMinutes.ToString
                dataFrame("timeTotal") = timeTotal.TotalMinutes.ToString
                dataFrame("hostName") = Net.Dns.GetHostName()

                '' Copy of "Ende.vb"

                '' Stimulus-Listen der Primingphasen im Datenfile ablegen
                'dataFrame("RF_practiceDL_primes") = rf1
                'dataFrame("RF_practiceDL_targets") = rf2
                'dataFrame("RF_practiceBH_primes") = rf3
                'dataFrame("RF_practiceBH_targets") = rf4
                'dataFrame("RF_targetsBH") = rf8
                'dataFrame("RF_primesDL") = rf5
                'dataFrame("RF_targetsDL") = rf6
                'dataFrame("RF_primesBH") = rf7

                saveCSV(dataFrame, "Data_RelAmb_" & Net.Dns.GetHostName & ".csv")

            Case Else
                Close()
        End Select
        instructionCount += 1
    End Sub





End Class
