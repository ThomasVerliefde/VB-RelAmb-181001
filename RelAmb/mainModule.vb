Imports System.IO

Module mainModule

    Public Class subjectPanel
        Inherits Panel

        Public Sub New(subjBox As TextBox, Optional fieldLeft As Integer = 250, Optional fieldTop As Integer = 250)

            Dim subjLabel As New Label()

            Location = New Point(fieldLeft, fieldTop)
            Size = New Size(200, 50)
            BorderStyle = BorderStyle.None

            subjBox.Location = New Point(85, 5)
            subjBox.Text = ""
            subjBox.Size = New Size(100, 40)
            subjBox.Font = New Font("Microsoft Sans Serif", 20)

            subjLabel.Location = New Point(0, 5)
            subjLabel.Size = New Size(100, 40)
            subjLabel.Text = "VPNr:"
            subjLabel.Font = New Font("Microsoft Sans Serif", 20)

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
            Font = New Font("Microsoft Sans Serif", 14)

        End Sub

    End Class

    Public Class instructionBox
        Inherits TextBox

        Public Sub New(Optional Txt As String = "", Optional horizontalDist As Double = 0.75, Optional verticalDist As Double = 0.67)

            Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

            Height = verticalDist * screenHeight
            Width = horizontalDist * screenWidth
            Text = Txt
            BackColor = Color.AliceBlue
            Multiline = True
            WordWrap = True
            Font = New Font("Microsoft Sans Serif", 22)
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


    Public Sub splitText(textFile As Object)
        Dim FinishedList As New List(Of String)
        Dim Lines = textFile.Split(" ")
        Dim temp As String = ""
        For Each line In Lines
            temp = line.Replace(vbLf, "")
            If Not String.IsNullOrWhiteSpace(temp) Then
                FinishedList.Add(temp)
            End If
        Next
        For Each item In FinishedList
            Console.WriteLine(item)
        Next
    End Sub

    Sub shuffleList(Of T)(list As IList(Of T))
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

    Public Function createPrimes(otherPrimes As List(Of String), posList As List(Of String), negList As List(Of String), strList As List(Of String))

        'otherNames = List of 4 names of SOs (2 pos & 2 neg)
        'posNouns = List of ALL positive noun primes (3L, 4L, ..., 10L)
        'negNouns = List of ALL negative noun primes (3L, 4L, ..., 10L)
        'neutStr = List of letter strings: repeats of 4 letters in 3L, 4L, ..., 10L (i.e. BBB, SSS, RRR, GGG, BBBB, SSSS, ..., GGGGGGGGGG)

        Dim valList As New List(Of List(Of String))({posList, negList})
        Dim Primes As New List(Of List(Of String))({New List(Of String)(4), New List(Of String)(4), New List(Of String)(4), New List(Of String)(4)})
        Dim nameL As Integer
        shuffleList(otherPrimes)

        For Each name In otherPrimes
            Dim index As Integer = otherPrimes.IndexOf(name)
            Dim index2 As Integer = index Mod 2
            Select Case name.Length
                Case <= 4
                    nameL = 1
                Case <= 6
                    nameL = 3
                Case <= 8
                    nameL = 5
                Case Else
                    nameL = 7
            End Select

            Primes(index2).Add(valList(index2)(nameL - 1 + Primes(index2).Count))
            Primes(2).Add(strList(index + (4 * (nameL - (name.Length Mod 2)))))
            Primes(3).Add(name)
        Next


        'For Each name In otherPrimes
        '    Dim index As Integer = otherPrimes.IndexOf(name)
        '    Select Case name.Length
        '        Case <= 4
        '            nameL = 0
        '        Case <= 6
        '            nameL = 2
        '        Case <= 8
        '            nameL = 4
        '        Case Else
        '            nameL = 6
        '    End Select
        '    If index Mod 2 = 0 Then
        '        posPrimes.Add(valList(0)(nameL + posPrimes.Count))
        '    Else
        '        negPrimes.Add(valList(1)(nameL + negPrimes.Count))
        '    End If
        '    strPrimes.Add(strList(index + (4 * (nameL - 3))))
        'Next

        'Primes = New List(Of List(Of String))({otherPrimes, posPrimes, negPrimes, strPrimes})
        Return Primes

    End Function

    'Public Function createTargets()

    '    'Select 2

    '    Return
    'End Function

    Public Function randomSelect(resource As Object)


        Return "Hello World"
    End Function

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
