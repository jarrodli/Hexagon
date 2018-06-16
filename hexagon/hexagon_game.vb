Public Class hexagon_game
    Public Current As Object = Nothing
    Public PTurn As Integer = 0
    Dim PicBox As PictureBox
    Dim STiles() As String = {"C1p1", "C1p2", "C1p3", "C2p1", "C2p2", "C2p3"}
    Public Player00 As New Clan(0, 0, STiles(0), 1)
    Public Player01 As New Clan(2, 0, STiles(1), 1)
    Public Player02 As New Clan(5, 0, STiles(2), 1)
    Public Player10 As New Clan(0, 15, STiles(3), 0)
    Public Player11 As New Clan(2, 15, STiles(4), 0)
    Public Player12 As New Clan(4, 15, STiles(5), 0)

    ' Creates a game and sets the board
    Private Sub hexagon_game_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AssignCoord(0, 5)
        AssignCoord(0, 6)
        AssignCoord(0, 8)
        AssignCoord(0, 9)
        AssignCoord(1, 5)
        AssignCoord(1, 6)
        AssignCoord(1, 7)
        AssignCoord(1, 8)
        AssignCoord(1, 9)
        AssignCoord(2, 5)
        AssignCoord(2, 6)
        AssignCoord(2, 7)
        AssignCoord(2, 8)
        AssignCoord(2, 9)
        AssignCoord(3, 5)
        AssignCoord(3, 7)
        AssignCoord(3, 8)
        AssignCoord(3, 9)
        AssignCoord(4, 5)
        AssignCoord(4, 6)
        AssignCoord(4, 7)
        AssignCoord(4, 8)
        AssignCoord(4, 9)
        AssignCoord(5, 5)
        AssignCoord(5, 6)
        AssignCoord(5, 7)
        AssignCoord(5, 8)
        AssignCoord(5, 9)
        UpdateScores(0)
        RandomizeGems()
    End Sub

    ' Randomize gems on board
    Public Sub RandomizeGems()

        Dim LegalGemHexs As Hex() = {Board(0, 5), Board(0, 6), Board(0, 8), Board(0, 9),
                                     Board(1, 5), Board(1, 6), Board(1, 7), Board(1, 8), Board(1, 9),
                                     Board(2, 5), Board(2, 6), Board(2, 7), Board(2, 8), Board(2, 9),
                                     Board(3, 5), Board(3, 7), Board(3, 8), Board(3, 9),
                                     Board(4, 6), Board(4, 7), Board(4, 8), Board(4, 9),
                                     Board(5, 5), Board(5, 6), Board(5, 7), Board(5, 8), Board(5, 9)}

        Dim LegalGemTiles As PictureBox() = {X0y5, X0y6, X0y8, X0y9,
                                             X1y5, X1y6, X1y7, X1y8, X1y9,
                                             X2y5, X2y6, X2y7, X2y8, X2y9,
                                             X3y5, X3y7, X3y8, X3y9,
                                             X4y6, X4y7, X4y8, X4y9,
                                             X5y5, X5y6, X5y7, X5y8, X5y9}

        For i As Integer = 0 To 2 Step 1
            Dim rand As Integer = GetRandom(0, 26)
            Dim pb As New PictureBox()

            While (LegalGemHexs(rand).occupied = True)
                rand = GetRandom(0, 26)
            End While

            Dim hexagon As Hex = LegalGemHexs(rand)

            hexagon.gem = 5
            hexagon.occupied = True

            LegalGemTiles(rand).Enabled = False

            With pb
                .Tag = hexagon
                .Height = 65
                .Width = 60
                .SizeMode = PictureBoxSizeMode.StretchImage
                .BackColor = Color.Transparent
                .Image = My.Resources._5
                .Name = "Gem5_" & i
                .Location = New Point(CInt(LegalGemTiles(rand).Left + LegalGemTiles(rand).Width - 65),
                                              CInt(LegalGemTiles(rand).Top + LegalGemTiles(rand).Height - 75))
            End With

            LegalGemHexs(rand).occupied = True
            LegalGemHexs(rand).gem = 5
            AddHandler pb.Click, AddressOf Gem_Click
            Controls.Add(pb)
            pb.BringToFront()
        Next

        For i As Integer = 0 To 4 Step 1
            Dim rand As Integer = GetRandom(0, 26)
            Dim pb As New PictureBox()

            While (LegalGemHexs(rand).occupied = True)
                rand = GetRandom(0, 26)
            End While

            Dim hexagon As Hex = LegalGemHexs(rand)

            hexagon.gem = 3
            hexagon.occupied = True

            LegalGemTiles(rand).Enabled = False

            With pb
                .Tag = hexagon
                .Height = 65
                .Width = 60
                .SizeMode = PictureBoxSizeMode.StretchImage
                .BackColor = Color.Transparent
                .Image = My.Resources._3
                .Name = "Gem3_" & i
                .Location = New Point(CInt(LegalGemTiles(rand).Left + LegalGemTiles(rand).Width - 65),
                                              CInt(LegalGemTiles(rand).Top + LegalGemTiles(rand).Height - 75))
            End With

            LegalGemHexs(rand).occupied = True
            LegalGemHexs(rand).gem = 3
            AddHandler pb.Click, AddressOf Gem_Click
            Controls.Add(pb)
            pb.BringToFront()
        Next

        Dim rand1 As Integer = GetRandom(0, 26)
        Dim pb1 As New PictureBox()

        While (LegalGemHexs(rand1).occupied = True)
            rand1 = GetRandom(0, 26)
        End While

        Dim hexagon1 As Hex = LegalGemHexs(rand1)

        hexagon1.gem = 9
        hexagon1.occupied = True

        LegalGemTiles(rand1).Enabled = False

        With pb1
            .Tag = hexagon1
            .Height = 65
            .Width = 60
            .SizeMode = PictureBoxSizeMode.StretchImage
            .BackColor = Color.Transparent
            .Image = My.Resources._9
            .Name = "Gem9_" & 1
            .Location = New Point(CInt(LegalGemTiles(rand1).Left + LegalGemTiles(rand1).Width - 65),
                                          CInt(LegalGemTiles(rand1).Top + LegalGemTiles(rand1).Height - 75))
        End With

        LegalGemHexs(rand1).occupied = True
        LegalGemHexs(rand1).gem = 9
        AddHandler pb1.Click, AddressOf Gem_Click
        Controls.Add(pb1)
        pb1.BringToFront()

    End Sub

    ' Plays movement sounds
    Sub PlayMovement()
        My.Computer.Audio.Play(My.Resources.movement, AudioPlayMode.Background)
    End Sub

    ' Plays capture sounds
    Sub PlayCapture()
        My.Computer.Audio.Play(My.Resources.movement, AudioPlayMode.WaitToComplete)
    End Sub

    'Plays gem capture sound
    Sub PlayCoin()
        My.Computer.Audio.Play(My.Resources.smw_coin, AudioPlayMode.WaitToComplete)
    End Sub

    ' Uses a static variable to allow random values based off of system clock 
    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Static Rand As System.Random = New System.Random()
        Return Rand.Next(Min, Max)
    End Function

    ' Updates a game when a move is made
    Public Sub MovePiece(Hexagon As Hex, NewCoord As Coord, Tile As PictureBox)
        If LegalMove(Board(NewCoord.x, NewCoord.y), NewCoord) Then
            Board(NewCoord.x, NewCoord.y).player = Current
            Board(NewCoord.x, NewCoord.y).piece = Current.GetTile
            Board(NewCoord.x, NewCoord.y).tag = Current.GetTag
            Board(Current.GetCoord.x, Current.GetCoord.y).occupied = False
            Current.SetCoord = NewCoord
            Board(NewCoord.x, NewCoord.y).occupied = True
            PicBox.Location = New Point(CInt(Tile.Left + Tile.Width - 75), CInt(Tile.Top + Tile.Height - 100))
            Moves -= 1
            lblMoves.Text = Moves
            PlayMovement()
        End If
    End Sub

    ' Assign a picturebox to a board
    Public Sub AssignCoord(x1 As Integer, y1 As Integer)
        Board(x1, y1).x = x1
        Board(x1, y1).y = y1
        Board(x1, y1).tile = DirectCast(Me.Controls.Find("X" & CStr(x1 & "y" & CStr(y1)), True).FirstOrDefault(), PictureBox)
    End Sub

    ' Determines if a player can move based off the rules of the game
    Public Function LegalMove(Hexagon As Hex, NewCoord As Coord) As Boolean
        Dim legal As Boolean = False

        If Current Is Nothing Or Moves = 0 Then
            Return legal
        End If
        Dim curx As Integer = Current.GetCoord.x
        Dim cury As Integer = Current.GetCoord.y
        Dim nx As Integer = NewCoord.x
        Dim ny As Integer = NewCoord.y
        If Current.GetCoord.x = NewCoord.x And Math.Abs(Current.GetCoord.y - NewCoord.y) = 1 Then
            legal = True
        ElseIf Current.GetCoord.y = NewCoord.y And Current.GetCoord.x - NewCoord.x < 0 And
            ((Current.GetCoord.x Mod 2 = 0 And Current.GetCoord.y Mod 2 = 1) Or
            (Current.GetCoord.x Mod 2 = 1 And Current.GetCoord.y Mod 2 = 0)) Then
            legal = True
        ElseIf Current.GetCoord.y = NewCoord.y And Current.GetCoord.x - NewCoord.x > 0 And
            ((Current.GetCoord.x Mod 2 = 1 And Current.GetCoord.y Mod 2 = 1) Or
            (Current.GetCoord.x Mod 2 = 0 And Current.GetCoord.y Mod 2 = 0)) Then
            legal = True
        End If
        If Hexagon.occupied = True And Hexagon.gem = 0 Then
            legal = False
        ElseIf Hexagon.occupied = True And Hexagon.gem <> 0 Then
            Current.SetScore = Hexagon.gem
            UpdateScores(Hexagon.gem)
        End If

        If (ny - cury) = 2 AndAlso Board(curx, cury + 1).tag <> Current.GetTag AndAlso Board(curx, cury + 1).occupied = True Then
            Dim nhex As Coord
            nhex.x = curx
            nhex.y = cury + 1
            TakeEnemy(nhex)
            legal = True
        ElseIf (ny - cury) = -2 AndAlso Board(curx, cury - 1).tag <> Current.GetTag AndAlso Board(curx, cury - 1).occupied = True Then
            Dim nhex As Coord
            nhex.x = curx
            nhex.y = cury - 1
            TakeEnemy(nhex)
            legal = True
        ElseIf (ny - cury) = 1 AndAlso Math.Abs(nx - curx) = 1 AndAlso Board(curx, cury + 1).tag <> Current.GetTag AndAlso
            Board(curx, cury + 1).occupied = True Then
            Dim nhex As Coord
            nhex.x = curx
            nhex.y = cury + 1
            TakeEnemy(nhex)
            legal = True
        ElseIf (ny - cury) = -1 AndAlso Math.Abs(nx - curx) = 1 AndAlso Board(curx, cury - 1).tag <> Current.GetTag AndAlso
            Board(curx, cury - 1).occupied = True Then
            Dim nhex As Coord
            nhex.x = curx
            nhex.y = cury - 1
            TakeEnemy(nhex)
            legal = True
        End If

        Return legal
    End Function

    Public Sub TakeEnemy(location As Coord)
        Dim pb As PictureBox = DirectCast(Me.Controls.Find(Board(location.x, location.y).piece, True).FirstOrDefault(), PictureBox)
        Current.SetScore = Board(location.x, location.y).player.GetScore
        UpdateScores(Board(location.x, location.y).player.GetScore)
        Me.Controls.Remove(pb)
        Board(location.x, location.y).occupied = False
    End Sub

    'Updates the game scores
    Public Sub UpdateScores(Value As Integer)
        If PTurn = 0 Then
            lblP1.Text = Player00.GetScore
            lblP2.Text = Player01.GetScore
            lblP3.Text = Player02.GetScore
            ScorePlayer1 += Value
            lblScore.Text = CStr(ScorePlayer1)
        Else
            lblP1.Text = Player10.GetScore
            lblP2.Text = Player11.GetScore
            lblP3.Text = Player12.GetScore
            ScorePlayer2 += Value
            lblScore.Text = CStr(ScorePlayer2)
        End If
    End Sub

    ' Sets the gamestate up for the next players turn
    Public Sub EndTurn()
        Moves = 0
        If PTurn = 0 Then
            PTurn = 1
        Else
            PTurn = 0
        End If
        UpdateScores(0)
        Current = Nothing
        Dice_Thrown = False
        Player = If(Player = 1, 2, 1)
        lblPlayer.Text = "Player " & Player
        lblMoves.Text = Moves
    End Sub

    '''                      '''
    ''' PLAYER CLICK EVENTS  '''
    '''                      '''

    Private Sub C1p1_Click(sender As Object, e As EventArgs) Handles C1p1.Click
        If PTurn = 0 Then
            Current = Player00
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    Private Sub C1p2_Click(sender As Object, e As EventArgs) Handles C1p2.Click
        If PTurn = 0 Then
            Current = Player01
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    Private Sub C1p3_Click(sender As Object, e As EventArgs) Handles C1p3.Click
        If PTurn = 0 Then
            Current = Player02
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    Private Sub C2p1_Click(sender As Object, e As EventArgs) Handles C2p1.Click
        If PTurn = 1 Then
            Current = Player10
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    Private Sub C2p2_Click(sender As Object, e As EventArgs) Handles C2p2.Click
        If PTurn = 1 Then
            Current = Player11
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    Private Sub C2p3_Click(sender As Object, e As EventArgs) Handles C2p3.Click
        If PTurn = 1 Then
            Current = Player12
            PicBox = DirectCast(sender, PictureBox)
        End If
    End Sub

    ' Handles a gem click event
    Private Sub Gem_Click(sender As Object, e As EventArgs)
        Dim THex As Coord
        Dim hexagon As Hex = sender.Tag
        THex.x = hexagon.x
        THex.y = hexagon.y
        If LegalMove(hexagon, THex) Then
            PlayCoin()
            hexagon.tile.Enabled = True
            MovePiece(hexagon, THex, hexagon.tile)
            sender.Image = Nothing
            sender.Dispose()
        End If
    End Sub

    '''                      '''
    ''' HEXAGON CLICK EVENTS '''
    '''                      '''

    Private Sub X0y0_Click(sender As Object, e As EventArgs) Handles X0y0.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y0
        xy.x = 0
        xy.y = 0
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y1_Click(sender As Object, e As EventArgs) Handles X0y1.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y1
        xy.x = 0
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y2_Click(sender As Object, e As EventArgs) Handles X0y2.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y2
        xy.x = 0
        xy.y = 2
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y3_Click(sender As Object, e As EventArgs) Handles X0y3.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y3
        xy.x = 0
        xy.y = 3
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y4_Click(sender As Object, e As EventArgs) Handles X0y4.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y4
        xy.x = 0
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y5Click(sender As Object, e As EventArgs) Handles X0y5.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y5
        xy.x = 0
        xy.y = 5
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y6_Click(sender As Object, e As EventArgs) Handles X0y6.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y6
        xy.x = 0
        xy.y = 6
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y8_Click(sender As Object, e As EventArgs) Handles X0y8.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y8
        xy.x = 0
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y9_Click(sender As Object, e As EventArgs) Handles X0y9.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y9
        xy.x = 0
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y10_Click(sender As Object, e As EventArgs) Handles X0y10.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y10
        xy.x = 0
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y12_Click(sender As Object, e As EventArgs) Handles X0y12.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y12
        xy.x = 0
        xy.y = 12
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y13_Click(sender As Object, e As EventArgs) Handles X0y13.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y13
        xy.x = 0
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y14_Click(sender As Object, e As EventArgs) Handles X0y14.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y14
        xy.x = 0
        xy.y = 14
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X0y15_Click(sender As Object, e As EventArgs) Handles X0y15.Click
        Dim xy As Coord
        Dim tile As PictureBox = X0y15
        xy.x = 0
        xy.y = 15
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y1_Click(sender As Object, e As EventArgs) Handles X1y1.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y1
        xy.x = 1
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y3_Click(sender As Object, e As EventArgs) Handles X1y3.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y3
        xy.x = 1
        xy.y = 3
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y4_Click(sender As Object, e As EventArgs) Handles X1y4.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y4
        xy.x = 1
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y5_Click(sender As Object, e As EventArgs) Handles X1y5.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y5
        xy.x = 1
        xy.y = 5
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y6_Click(sender As Object, e As EventArgs) Handles X1y6.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y6
        xy.x = 1
        xy.y = 6
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y7_Click(sender As Object, e As EventArgs) Handles X1y7.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y7
        xy.x = 1
        xy.y = 7
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y8_Click(sender As Object, e As EventArgs) Handles X1y8.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y8
        xy.x = 1
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y9_Click(sender As Object, e As EventArgs) Handles X1y9.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y9
        xy.x = 1
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y10_Click(sender As Object, e As EventArgs) Handles X1y10.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y10
        xy.x = 1
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y11_Click(sender As Object, e As EventArgs) Handles X1y11.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y11
        xy.x = 1
        xy.y = 11
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y13_Click(sender As Object, e As EventArgs) Handles X1y13.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y13
        xy.x = 1
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y14_Click(sender As Object, e As EventArgs) Handles X1y14.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y14
        xy.x = 1
        xy.y = 14
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X1y15_Click(sender As Object, e As EventArgs) Handles X1y15.Click
        Dim xy As Coord
        Dim tile As PictureBox = X1y15
        xy.x = 1
        xy.y = 15
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y0_Click(sender As Object, e As EventArgs) Handles X2y0.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y0
        xy.x = 2
        xy.y = 0
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y1_Click(sender As Object, e As EventArgs) Handles X2y1.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y1
        xy.x = 2
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, tile)

    End Sub

    Private Sub X2y2_Click(sender As Object, e As EventArgs) Handles X2y2.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y2
        xy.x = 2
        xy.y = 2
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y3_Click(sender As Object, e As EventArgs) Handles X2y3.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y3
        xy.x = 2
        xy.y = 3
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y4_Click(sender As Object, e As EventArgs) Handles X2y4.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y4
        xy.x = 2
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y5_Click(sender As Object, e As EventArgs) Handles X2y5.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y5
        xy.x = 2
        xy.y = 5
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y6_Click(sender As Object, e As EventArgs) Handles X2y6.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y6
        xy.x = 2
        xy.y = 6
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y7_Click(sender As Object, e As EventArgs) Handles X2y7.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y7
        xy.x = 2
        xy.y = 7
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y8_Click(sender As Object, e As EventArgs) Handles X2y8.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y8
        xy.x = 2
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y9_Click(sender As Object, e As EventArgs) Handles X2y9.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y9
        xy.x = 2
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y10_Click(sender As Object, e As EventArgs) Handles X2y10.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y10
        xy.x = 2
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y11_Click(sender As Object, e As EventArgs) Handles X2y11.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y11
        xy.x = 2
        xy.y = 11
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y12_Click(sender As Object, e As EventArgs) Handles X2y12.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y12
        xy.x = 2
        xy.y = 12
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y13_Click(sender As Object, e As EventArgs) Handles X2y13.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y13
        xy.x = 2
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y14_Click(sender As Object, e As EventArgs) Handles X2y14.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y14
        xy.x = 2
        xy.y = 14
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X2y15_Click(sender As Object, e As EventArgs) Handles X2y15.Click
        Dim xy As Coord
        Dim tile As PictureBox = X2y15
        xy.x = 2
        xy.y = 15
        MovePiece(Board(xy.x, xy.y), xy, tile)
    End Sub

    Private Sub X3y10_Click(sender As Object, e As EventArgs) Handles X3y10.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y13_Click(sender As Object, e As EventArgs) Handles X3y13.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y14_Click(sender As Object, e As EventArgs) Handles X3y14.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 14
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y12_Click(sender As Object, e As EventArgs) Handles X3y12.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 12
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y11_Click(sender As Object, e As EventArgs) Handles X3y11.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 11
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y8_Click(sender As Object, e As EventArgs) Handles X3y8.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y9_Click(sender As Object, e As EventArgs) Handles X3y9.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y7_Click(sender As Object, e As EventArgs) Handles X3y7.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 7
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y4_Click(sender As Object, e As EventArgs) Handles X3y4.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y5_Click(sender As Object, e As EventArgs) Handles X3y5.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 5
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y3_Click(sender As Object, e As EventArgs) Handles X3y3.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 3
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y2_Click(sender As Object, e As EventArgs) Handles X3y2.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 2
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y1_Click(sender As Object, e As EventArgs) Handles X3y1.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X3y0_Click(sender As Object, e As EventArgs) Handles X3y0.Click
        Dim xy As Coord
        xy.x = 3
        xy.y = 0
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y0_Click(sender As Object, e As EventArgs) Handles X4y0.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 0
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y1_Click(sender As Object, e As EventArgs) Handles X4y1.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y2_Click(sender As Object, e As EventArgs) Handles X4y2.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 2
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y3_Click(sender As Object, e As EventArgs) Handles X4y3.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 3
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y4_Click(sender As Object, e As EventArgs) Handles X4y4.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y6_Click(sender As Object, e As EventArgs) Handles X4y6.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 6
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y7_Click(sender As Object, e As EventArgs) Handles X4y7.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 7
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y8_Click(sender As Object, e As EventArgs) Handles X4y8.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y9_Click(sender As Object, e As EventArgs) Handles X4y9.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y10_Click(sender As Object, e As EventArgs) Handles X4y10.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y11_Click(sender As Object, e As EventArgs) Handles X4y11.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 11
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y13_Click(sender As Object, e As EventArgs) Handles X4y13.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X4y15_Click(sender As Object, e As EventArgs) Handles X4y15.Click
        Dim xy As Coord
        xy.x = 4
        xy.y = 15
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y0_Click(sender As Object, e As EventArgs) Handles X5y0.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 0
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y1_Click(sender As Object, e As EventArgs) Handles X5y1.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 1
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y2_Click(sender As Object, e As EventArgs) Handles X5y2.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 2
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y4_Click(sender As Object, e As EventArgs) Handles X5y4.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 4
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y5_Click(sender As Object, e As EventArgs) Handles X5y5.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 5
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y6_Click(sender As Object, e As EventArgs) Handles X5y6.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 6
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y7_Click(sender As Object, e As EventArgs) Handles X5y7.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 7
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y8_Click(sender As Object, e As EventArgs) Handles X5y8.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 8
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y9_Click(sender As Object, e As EventArgs) Handles X5y9.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 9
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y10_Click(sender As Object, e As EventArgs) Handles X5y10.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 10
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y11_Click(sender As Object, e As EventArgs) Handles X5y11.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 11
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y12_Click(sender As Object, e As EventArgs) Handles X5y12.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 12
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y13_Click(sender As Object, e As EventArgs) Handles X5y13.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 13
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y14_Click(sender As Object, e As EventArgs) Handles X5y14.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 14
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub X5y15_Click(sender As Object, e As EventArgs) Handles X5y15.Click
        Dim xy As Coord
        xy.x = 5
        xy.y = 15
        MovePiece(Board(xy.x, xy.y), xy, sender)
    End Sub

    Private Sub btnThrowDice_Click(sender As Object, e As EventArgs) Handles btnThrowDice.Click
        If Dice_Thrown = False Then
            d1.Visible = True
            d1.BringToFront()
            Timer1.Start()
            Dice_Thrown = True
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar.Increment(5)
        If ProgressBar.Value = 100 Then
            Timer1.Stop()
            ProgressBar.Value = 0
            d1.Visible = False
            d2.Image = d1.Image
            lblMoves.Text = Moves
        Else
            Moves = GetRandom(1, 6)
            If Moves = 1 Then
                d1.Image = My.Resources.DiceRoll1
            ElseIf Moves = 2 Then
                d1.Image = My.Resources.DiceRoll2
            ElseIf Moves = 3 Then
                d1.Image = My.Resources.DiceRoll3
            ElseIf Moves = 4 Then
                d1.Image = My.Resources.DiceRoll4
            ElseIf Moves = 5 Then
                d1.Image = My.Resources.DiceRoll5
            ElseIf Moves = 6 Then
                d1.Image = My.Resources.DiceRoll6
            End If
        End If
    End Sub

    ' remove this at the end
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Moves += 100
    End Sub

    Private Sub btnEndTurn_Click(sender As Object, e As EventArgs) Handles btnEndTurn.Click
        EndTurn()
    End Sub

    '  Changes colour of picturebox when Leaveed over
    Private Sub PictureBox1_MouseEnter(sender As Object, e As System.EventArgs) Handles X5y9.MouseEnter, X5y8.MouseEnter, X5y7.MouseEnter, X5y6.MouseEnter,
            X5y5.MouseEnter, X5y4.MouseEnter, X5y2.MouseEnter, X5y15.MouseEnter, X5y14.MouseEnter, X5y13.MouseEnter, X5y12.MouseEnter, X5y11.MouseEnter,
        X5y10.MouseEnter, X5y1.MouseEnter, X5y0.MouseEnter, X4y9.MouseEnter, X4y8.MouseEnter, X4y7.MouseEnter, X4y6.MouseEnter, X4y4.MouseEnter,
        X4y3.MouseEnter, X4y2.MouseEnter, X4y15.MouseEnter, X4y13.MouseEnter, X4y11.MouseEnter, X4y10.MouseEnter, X4y1.MouseEnter, X4y0.MouseEnter,
        X3y13.MouseEnter, X3y12.MouseEnter, X3y11.MouseEnter, X3y10.MouseEnter, X3y1.MouseEnter, X3y0.MouseEnter, X2y9.MouseEnter, X2y8.MouseEnter,
        X2y7.MouseEnter, X2y6.MouseEnter, X2y5.MouseEnter, X2y4.MouseEnter, X2y3.MouseEnter, X2y2.MouseEnter, X2y15.MouseEnter, X2y14.MouseEnter,
        X2y13.MouseEnter, X2y12.MouseEnter, X2y11.MouseEnter, X2y10.MouseEnter, X2y1.MouseEnter, X2y0.MouseEnter, X1y9.MouseEnter, X1y8.MouseEnter,
        X1y7.MouseEnter, X1y6.MouseEnter, X1y5.MouseEnter, X1y4.MouseEnter, X1y3.MouseEnter, X1y15.MouseEnter, X1y14.MouseEnter, X1y13.MouseEnter,
        X1y11.MouseEnter, X1y10.MouseEnter, X1y1.MouseEnter, X0y9.MouseEnter, X0y8.MouseEnter, X0y6.MouseEnter, X0y5.MouseEnter, X0y4.MouseEnter,
        X0y3.MouseEnter, X0y2.MouseEnter, X0y15.MouseEnter, X0y14.MouseEnter, X0y13.MouseEnter, X0y12.MouseEnter, X0y10.MouseEnter, X0y1.MouseEnter,
        X0y0.MouseEnter, X3y9.MouseEnter, X3y8.MouseEnter, X3y7.MouseEnter, X3y5.MouseEnter, X3y4.MouseEnter, X3y3.MouseEnter, X3y2.MouseEnter
        sender.BackColor = Color.PaleGreen
    End Sub

    'Changes colour of pictrebox when not Leaveed over 
    Private Sub PictureBox1_MouseLeave(sender As Object, e As System.EventArgs) Handles X5y9.MouseLeave, X5y8.MouseLeave, X5y7.MouseLeave, X5y6.MouseLeave,
        X5y5.MouseLeave, X5y4.MouseLeave, X5y2.MouseLeave, X5y15.MouseLeave, X5y14.MouseLeave, X5y13.MouseLeave, X5y12.MouseLeave, X5y11.MouseLeave,
        X5y10.MouseLeave, X5y1.MouseLeave, X5y0.MouseLeave, X4y9.MouseLeave, X4y8.MouseLeave, X4y7.MouseLeave, X4y6.MouseLeave, X4y4.MouseLeave,
        X4y3.MouseLeave, X4y2.MouseLeave, X4y15.MouseLeave, X4y13.MouseLeave, X4y11.MouseLeave, X4y10.MouseLeave, X4y1.MouseLeave, X4y0.MouseLeave,
        X3y9.MouseLeave, X3y8.MouseLeave, X3y7.MouseLeave, X3y5.MouseLeave, X3y4.MouseLeave, X3y3.MouseLeave, X3y2.MouseLeave, X3y14.MouseLeave,
        X3y13.MouseLeave, X3y12.MouseLeave, X3y11.MouseLeave, X3y10.MouseLeave, X3y1.MouseLeave, X3y0.MouseLeave, X2y9.MouseLeave, X2y8.MouseLeave,
        X2y7.MouseLeave, X2y6.MouseLeave, X2y5.MouseLeave, X2y4.MouseLeave, X2y3.MouseLeave, X2y2.MouseLeave, X2y15.MouseLeave, X2y14.MouseLeave,
        X2y13.MouseLeave, X2y12.MouseLeave, X2y11.MouseLeave, X2y10.MouseLeave, X2y1.MouseLeave, X2y0.MouseLeave, X1y9.MouseLeave, X1y8.MouseLeave,
        X1y7.MouseLeave, X1y6.MouseLeave, X1y5.MouseLeave, X1y4.MouseLeave, X1y3.MouseLeave, X1y15.MouseLeave, X1y14.MouseLeave, X1y13.MouseLeave,
        X1y11.MouseLeave, X1y10.MouseLeave, X1y1.MouseLeave, X0y9.MouseLeave, X0y8.MouseLeave, X0y6.MouseLeave, X0y5.MouseLeave, X0y4.MouseLeave,
        X0y3.MouseLeave, X0y2.MouseLeave, X0y15.MouseLeave, X0y14.MouseLeave, X0y13.MouseLeave, X0y12.MouseLeave, X0y10.MouseLeave, X0y1.MouseLeave,
        X0y0.MouseLeave
        sender.BackColor = Color.Transparent
    End Sub
End Class