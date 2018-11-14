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
