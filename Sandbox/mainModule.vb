Using RDotNet

Module Module1

    Public Class continueButton
        Inherits Button

        Public Sub New()

            Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

            Height = 40
            Width = 150
            Text = "Weiter"
            Font = New Font("Microsoft Sans Serif", 14)
            Top = (4 * screenHeight / 5) - (Me.Height / 2)
            Left = (screenWidth / 2) - (Me.Width / 2)

        End Sub

    End Class

    Public Function blockRand(ByVal N As Integer)

        Dim rnd As New Random(20181001)
















        Return Array
    End Function

    Sub mainTest()





        Dim foo As List(Of List(Of Integer))

            foo = blockRand(1)

            For i As Integer = 0 To foo.Count - 1

                Console.WriteLine(String.Join(", ", foo(i)))

            Next

    End Sub



End Module
