﻿Imports NodaTime

Public Class mainForm
	Inherits Form

	Public instructionCount As New Integer 'main counter, to control the flow of the experiment; increases after each continuebutton press in this form

	Friend WithEvents contButton As New continueButton 'main button, to advance the flow of the experiment
	Private ReadOnly instrText As New instructionBox 'main method of displaying the instructions to participants; disabled & readonly

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
	Friend firstOthers As String

	Friend otherPos As New List(Of String)
	Friend otherNeg As New List(Of String)



	'   '####################################'
	'   'Imported CODE - Not Yet Consolidated'

	'   'Variablen für die Übungsphasen

	'   Public practiceBH = New String(7, 3) {} ' Berger & Hütter Paradigma

	'   ' Aufgespaltete RF für Datenspeicherung

	'   Public practiceBHprimes = New String(7) {}
	'   Public practiceBHtargets = New String(7) {}

	'   Public practiceposadj = New String() {"geduldig", "zärtlich", "humorvoll", "fleissig"}
	'   Public practicenegadj = New String() {"boshaft", "korrupt", "ungerecht", "gehässig"}

	'   Public practiceposnoun = New String() {"Kuss", "Idee", "Chance"}
	'   Public practicenegnoun = New String() {"Mord", "Angst", "Sklave"}
	'   Public practiceambnoun = New String() {"Arbeit", "Macht", "Diät"}


	'   ' Variablen für affektives Priming deLiver/ Berger & Hütter

	'   ' Wortmaterial Nomen
	'   Public nounspos = New String() {"Mut", "Lust", "Glück", "Freude"}
	'   Public nounsneg = New String() {"Leid", "Ärger", "Furcht", "Trauer"}
	'   Public nounsnonwords = New String() {"hdreeh", "izmcsg", "tdkesu", "unerit", "ilbuel", "rdnslb", "ursdcu", "nsedni"}
	'   Public nounsamb = New String() {"Karriere", "Regulierung", "Arztbesuch", "Protest", "Alkohol", "Feuer", "Smartphone", "Stolz"}

	'   ' Wortmaterial Adjektive
	'   Public adjpos = New String() {"frei", "klug", "treu", "gesund", "beliebt", "ehrlich", "herzlich", "friedlich"}
	'   Public adjneg = New String() {"blöd", "dumm", "fies", "brutal", "grausam", "neidisch", "peinlich", "widerlich"}
	'   Public adjnonwords = New String() {"takatg", "ineifk", "aahrte", "ndwrow", "rurtum", "worenl", "iasweg", "enoers"}

	'   ' Liste mit allen Nomen für die expliziten Ratings
	'   ' Keine Nonwords!
	'   Public nouns = New String(15, 1) {}




	'   ' Material Berger & Hütter (96 Trials)
	'   ' PRIMES = NOUNS: amb (8), neutral (8), pos (4), neg (4)
	'   ' TARGETS = ADJECTIVES: pos (8), neg (8)
	'   ' 24 Primes * 2 Target-Types = 48 Trials * 2 (höhere Reliabilität)
	'   Public trialsBH = New String(95, 3) {}
	'   Public primesBH = New String(95) {}
	'   Public targetsBH = New String(95) {}

	'   Public rf1 As String ' Practice trials primes
	'   Public rf2 As String ' Practice trials targets
	'   Public rf3 As String ' Test trials primes
	'   Public rf4 As String ' Test trials targets
	'   Public rf5 As String ' ...
	'   Public rf6 As String
	'   Public rf7 As String
	'   Public rf8 As String

	''########################################################'




	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		subjectForm.ShowDialog()

		Me.Controls.Add(Me.instrText)
		xCenter(Me.instrText, 0.4)
		Me.instrText.Text = My.Resources.ResourceManager.GetString("_0_mainInstr")

		Me.Controls.Add(Me.contButton)
		xCenter(Me.contButton)

	End Sub

	Public Sub loadNext(sender As Object, e As EventArgs) Handles contButton.Click
		Select Case Me.instructionCount
			Case 0 'Start of Experiment
				subjectForm.Dispose()
				Me.startT = time.GetCurrentInstant()
				Me.instrText.Text = My.Resources.ResourceManager.GetString("_1_otherInstr")




			Case 1 'Collecting Names of 'Significant Others'
				Me.otherT = time.GetCurrentInstant()
				otherForm.ShowDialog()
				otherForm.Dispose()

				' Create All Practice & Experiment Trials
				' For the Practice trials, get a set of X*2 othernames (removing othernames already suggested by the participant, then randomly choosing)

				' Practice Trials

				' NOT FINISHED !!!! -> WE NEED ^^



				Dim otherPractice As New List(Of String)(My.Resources.practiceOthers.Split(" "))
				otherPractice = compareList(Me.otherNeg, otherPractice)
				otherPractice = compareList(Me.otherPos, otherPractice)
				shuffleList(otherPractice)
				Dim practicePrime_Pos As New List(Of String)(My.Resources.practicePrime_Pos.Split(" "))
				shuffleList(practicePrime_Pos)
				Dim practicePrime_Neg As New List(Of String)(My.Resources.practicePrime_Neg.Split(" "))
				shuffleList(practicePrime_Neg)
				Dim practicePrime_Str As New List(Of String)(My.Resources.practicePrime_Str.Split(" "))
				shuffleList(practicePrime_Str)

				'practicePrimes is a List of List of String with 5 lists: PositiveNounPrimes, NegativeNounPrimes, PositiveOthers, NegativeOthers, MatchingLetterStrings
				Dim practicePrimes As New List(Of List(Of String))({
								New List(Of String)({otherPractice(0)}),
								New List(Of String)({otherPractice(1)}),
								New List(Of String)({practicePrime_Pos(0)}),
								New List(Of String)({practicePrime_Neg(0)}),
								New List(Of String)({practicePrime_Str(0)})
								})

				Dim practiceTrials = createTrials(
										practicePrimes,
										New List(Of List(Of String))({
											 New List(Of String)(My.Resources.practiceTarget_Pos.Split(" ")),
											 New List(Of String)(My.Resources.practiceTarget_Neg.Split(" "))
											 }),
										timesPrimes:=2 'How often each prime is paired with a target from each category
										)
				' Results in 20 Trials (Can shorten to 10 by setting timesPrimes to 1)

				''Debug
				'For Each list In practiceTrials
				'	For Each item In list
				'		Console.Write(item & (" "))
				'	Next
				'	Console.WriteLine("")
				'Next
				''End Debug

				' Experiment Trials

				Dim experimentPrimes = createPrimes(
					Me.otherPos,
					Me.otherNeg,
					New List(Of String)(My.Resources.experimentPrime_Pos.Split(" ")),
					New List(Of String)(My.Resources.experimentPrime_Neg.Split(" ")),
					New List(Of String)(My.Resources.experimentPrime_Str.Split(" "))
					)

				Dim experimentTrials = createTrials(
										experimentPrimes,
										New List(Of List(Of String))({
											 New List(Of String)(My.Resources.experimentTarget_Pos.Split(" ")),
											 New List(Of String)(My.Resources.experimentTarget_Neg.Split(" "))
											 }),
										timesPrimes:=4 'How often each prime is paired with a target from each category
										)
				' Results in 96 Trials

				'Debug
				Console.WriteLine("---------------------")
				Dim amount As Integer
				For Each c In experimentTrials
					For Each d In c
						Console.Write(" * " + d)
					Next
					amount += c.Count
					Console.WriteLine("")
				Next
				Console.WriteLine("")
				'End Debug

				Me.instrText.Text = My.Resources.ResourceManager.GetString("_2_practice" & Me.keyAss)
			Case 2 'Practice Trials





				Me.practiceT = time.GetCurrentInstant()
				'practiceForm.ShowDialog()

				Me.instrText.Text = My.Resources.ResourceManager.GetString("_3_experiment" & Me.keyAss)
			Case 3 'Experiment Proper
				Me.experimentT = time.GetCurrentInstant()
				'expForm.ShowDialog()

				Me.instrText.Text = My.Resources.ResourceManager.GetString("_4_explicitInstr")
			Case 4 'Explicit Measurements of Ambivalence
				Me.explicitT = time.GetCurrentInstant()
				explicitForm.ShowDialog()

				Me.instrText.Text = My.Resources.ResourceManager.GetString("_5_demoInstr")
			Case 5 'Demographic Information
				explicitForm.Dispose()
				Me.demographicsT = time.GetCurrentInstant()
				demographicsForm.ShowDialog()

			Case 6 'End of Experiment
				demographicsForm.Dispose()

				Me.instrText.Text = My.Resources.ResourceManager.GetString("_6_endInstr")
				Me.instrText.Font = New Font("Microsoft Sans Serif", 40)
				'instrText.TextAlign = HorizontalAlignment.Center
				Me.contButton.Text = "Abbrechen"
				Me.endT = time.GetCurrentInstant()

				Me.timeOther = Me.practiceT - Me.otherT
				Me.timePractice = Me.experimentT - Me.practiceT
				Me.timeExperiment = Me.explicitT - Me.experimentT
				Me.timeExplicit = Me.demographicsT - Me.explicitT
				Me.timeDemographics = Me.endT - Me.demographicsT
				Me.timeTotal = Me.endT - Me.startT

				dataFrame("timeOther") = Me.timeOther.TotalMinutes.ToString
				dataFrame("timePractice") = Me.timePractice.TotalMinutes.ToString
				dataFrame("timeExperiment") = Me.timeExperiment.TotalMinutes.ToString
				dataFrame("timeExplicit") = Me.timeExplicit.TotalMinutes.ToString
				dataFrame("timeDemographics") = Me.timeDemographics.TotalMinutes.ToString
				dataFrame("timeTotal") = Me.timeTotal.TotalMinutes.ToString
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
				Me.Close()
		End Select
		Me.instructionCount += 1
	End Sub





End Class
