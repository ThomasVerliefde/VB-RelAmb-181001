Imports System.IO
Imports System.Text.RegularExpressions

Module mainModule

    Friend sansSerif22 = New Font("Microsoft Sans Serif", 22)
    Friend sansSerif20 = New Font("Microsoft Sans Serif", 20)
    Friend sansSerif14 = New Font("Microsoft Sans Serif", 14)

    Public Class labelledBox
        Inherits Panel

        Public Sub New(subjBox As TextBox, Optional labelText As String = "VPNr:", Optional boxWidth As Integer = 100, Optional fieldHeight As Integer = 50,
                Optional fieldLeft As Integer = 250, Optional fieldTop As Integer = 250)

            Dim subjLabel As New Label()
            Dim labelWidth = TextRenderer.MeasureText(labelText, sansSerif20).Width

            Location = New Point(fieldLeft, fieldTop)
            Size = New Size((labelWidth + boxWidth) * 1.1, fieldHeight)
            BorderStyle = BorderStyle.None

            subjBox.Location = New Point(labelWidth + 10, fieldHeight * 0.1)
            subjBox.Text = ""
            subjBox.Size = New Size(boxWidth, fieldHeight)
            subjBox.Font = sansSerif20

			subjLabel.Location = New Point(0, fieldHeight * 0.1)
			subjLabel.Size = New Size(labelWidth, fieldHeight)
			subjLabel.Text = labelText
			subjLabel.Font = sansSerif20

			Controls.Add(subjBox)
            Controls.Add(subjLabel)

        End Sub

    End Class

    Public Class continueButton
        Inherits Button

        Public Sub New(Optional Txt As String = "Weiter", Optional buttonHeight As Integer = 40, Optional buttonWidth As Integer = 150)

            Height = buttonHeight
            Width = buttonWidth
            Text = Txt
            Font = sansSerif14

        End Sub

    End Class

    Public Class instructionBox
        Inherits RichTextBox

        Public Sub New(Optional Txt As String = "", Optional horizontalDist As Double = 0.75, Optional verticalDist As Double = 0.67)

            Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

            Height = verticalDist * screenHeight
            Width = horizontalDist * screenWidth
            Text = Txt
            BackColor = Color.AliceBlue
            Multiline = True
            WordWrap = True
            Font = sansSerif22
            [ReadOnly] = True
            Enabled = False

        End Sub

    End Class


    Public Sub xCenter(any As Object, Optional verticalDist As Double = 0.85, Optional horizontalDist As Double = 0.5, Optional nudgeLeft As Integer = 0, Optional nudgeTop As Integer = 0)

        Dim formWidth As Integer = Form.ActiveForm.DesktopBounds.Width
        Dim formHeight As Integer = Form.ActiveForm.DesktopBounds.Height

        Try
            any.Left = (horizontalDist * formWidth) - (any.Width / 2) - nudgeLeft
            any.Top = (verticalDist * formHeight) - (any.Height / 2) - nudgeTop
        Catch ex As Exception
        End Try

    End Sub

    Public Function setCond(subjN As Integer)
        Return Val(My.Resources.BlockRandomisation((subjN - 1) * 2))
    End Function

#Disable Warning IDE1006 ' Naming Styles
	Public Function IsName(ByVal checkString As String)
#Enable Warning IDE1006 ' Naming Styles

		' Old Code, does basically the same as the first Regex function
		'For i = 0 To checkString.Length - 1
		'          If Not Char.IsLetter(checkString.Chars(i)) Then
		'              Return False
		'          End If
		'      Next

		Return If(Regex.IsMatch(checkString, "^([a-zA-ZÀ-ÿ])*$") OrElse Regex.IsMatch(checkString, "^(([a-zA-ZÀ-ÿ]){1,}(-){0,1}([a-zA-ZÀ-ÿ])*){0,}$"),
			True,
			False)
	End Function


	Public Sub shuffleList(Of T)(list As IList(Of T)) 'Simple shuffle sub, switching around items without constraints
        Dim r As Random = New Random()
        For i = 0 To list.Count - 1
            Dim index As Integer = r.Next(i, list.Count)
            If i <> index Then
                ' swap list(i) and list(index)
                Dim temp As T = list(i)
                list(i) = list(index)
                list(index) = temp
            End If
        Next
    End Sub

    Public Function createPrimes(otherPos As List(Of String), otherNeg As List(Of String), posList As List(Of String), negList As List(Of String), strList As List(Of String))

		'otherNames = List of 4 names of SOs (2 pos & 2 neg)
		'posList = List of ALL positive noun primes (3L, 4L, ..., 10L)
		'negList = List of ALL negative noun primes (3L, 4L, ..., 10L)
		'strList = List of letter strings: repeats of 4 letters in 3L, 4L, ..., 10L (i.e. BBB, SSS, RRR, GGG, BBBB, SSSS, ..., GGGGGGGGGG)

		Dim otherPrimes As New List(Of String)({otherPos.Concat(otherNeg).ToString()})
		Dim valList As New List(Of List(Of String))({posList, negList})
        Dim Primes As New List(Of List(Of String))({New List(Of String)(2), New List(Of String)(2), New List(Of String)(4), New List(Of String)(4)})
        'shuffleList(otherPrimes) 'I'm not really sure whether it's better to shuffle these names or not.
        'Assuming the order is (Pos, Pos, Neg, Neg), length pairings will be crossed (posName + posPrime, posName + negPrime, negName + posPrime, negName + negPrime)

        For Each name In otherPrimes
            Dim index As Integer = otherPrimes.IndexOf(name)
            Dim index2 As Integer = index Mod 2
            Dim nameL As Integer = Math.Max(Math.Min(name.Length - 3 + name.Length Mod 2, 7), 1)
            'Bins the length of otherName to 3-4, 5-6, 7-8, and 9-10, transcribed as 1, 3, 5, & 7
            'Takes into account that names could be shorter than 3 or longer than 10, and makes sure the result is maximum 7 and minimum 1
            Primes(index2).Add(valList(index2)(nameL - 1 + Primes(index2).Count))
            'Fills either Primes(0) or Primes(1), with valList(0) or valList(1) items (0 will be positive, 1 will be negative)
            'The selected item in valList is matched to be similarly long as the otherName
            'For the first positive and negative item, within bins the shorter word is chosen; opposite is true for the second items
            'This should prevent words being used twice
            ' Another way to tackle this, could be to set wordlists with 2 words for each length, and match exactly on length, but would require quite some recoding
            Primes(2).Add(strList(index + (4 * (nameL - (name.Length Mod 2)))))
            'Adds a letter string, exactly matching the length of otherName
            '4 different letters are used (B,S,G,R) and as such, 1 match for each otherName
            Primes(3).Add(name)
            'Also adds the otherName
        Next
        Return Primes

    End Function

    'Public Function createTargets()

    '    'Select 2

    '    Return
    'End Function

    Public Function randomSelect(resource As Object)


        Return "Hello World"
    End Function

	'Does createTrials one work? 

	Public Function createTrials(Primes As List(Of List(Of String)), Target As List(Of List(Of String)))

        Dim numP As Integer = Primes.Count
        Dim numT As Integer = Target.Count
        Dim Trials = New List(Of List(Of String))

        For i = 0 To numP - 1
            For j = 0 To numT - 1
                For Each itemP In Primes(i)
                    For Each itemT In Target(j)
                        Trials.Add(New List(Of String)({itemP, itemT, i.ToString, j.ToString}))
                    Next
                Next
            Next
        Next

        Return Trials
    End Function

    Public Sub saveCSV(ByVal dataFrame As Dictionary(Of String, String), Optional ByVal path As String = "rawData.csv")

        Dim fileOutput As New String("")

        If Not My.Computer.FileSystem.FileExists(path) Then

            Dim colNames As New String("")

            Try
                'When there is no dataframe yet, first create a headers row
                For Each key In dataFrame.Keys
                    colNames += key + ","
                Next
                colNames = colNames.Substring(0, colNames.Length - 1) + Environment.NewLine

                File.Create(path).Dispose()
                My.Computer.FileSystem.WriteAllText(path, colNames, True)

            Catch ex As Exception
                MessageBox.Show("Beim Erstellen der Datei " + path + " ist ein Fehler aufgetreten: \n" + ex.ToString)
            End Try
        End If

        'Putting the dictionary data in the CSV format
        For Each value In dataFrame.Values
            fileOutput += value + ","
        Next
        fileOutput = fileOutput.Substring(0, fileOutput.Length - 1) + Environment.NewLine

        'Saving the data
        Try
            My.Computer.FileSystem.WriteAllText(path, fileOutput, True)

        Catch ex As Exception
            MessageBox.Show("Beim Speichern der Datei" + path + " ist ein Fehler aufgetreten: \n" + ex.ToString)
        End Try

    End Sub

End Module
