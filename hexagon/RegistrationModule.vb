Public Class RegistrationModule

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' My.Computer.Audio.Play(My.Resources., AudioPlayMode.Background)
        Player1Name = tbC1.Text
        Player2Name = tbC2.Text
        Dim box = New hexagon_game()
        box.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class