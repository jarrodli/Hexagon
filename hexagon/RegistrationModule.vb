Public Class RegistrationModule
    Public Shared Player1Name As String
    Public Shared Player2Name As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' My.Computer.Audio.Play(My.Resources., AudioPlayMode.Background)
        Player1Name = tbC1.Text
        Player2Name = tbC2.Text
        Dim box = New hexagon_game()
        box.Show()
        Me.Close()
    End Sub
End Class