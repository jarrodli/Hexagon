''' 
''' Form1.vb
''' Jarrod Li
''' Hexagon Board Game
''' 7 May 2018
'''


Public Class Form1

    Public Current As Object = Nothing
    Public PTurn As Integer = 0
    Dim PicBox As PictureBox
    Dim STiles() As PictureBox = {DirectCast(Me.Controls.Find("X0y0", True).FirstOrDefault(), PictureBox),
                                  DirectCast(Me.Controls.Find("X2y0", True).FirstOrDefault(), PictureBox),
                                  DirectCast(Me.Controls.Find("X5y0", True).FirstOrDefault(), PictureBox),
                                  DirectCast(Me.Controls.Find("X0y15", True).FirstOrDefault(), PictureBox),
                                  DirectCast(Me.Controls.Find("X2y15", True).FirstOrDefault(), PictureBox),
                                  DirectCast(Me.Controls.Find("X4y15", True).FirstOrDefault(), PictureBox)}
    Public Player00 As New Clan(C1p1, 0, 0, STiles(0))
    Public Player01 As New Clan(C1p2, 2, 0, STiles(1))
    Public Player02 As New Clan(C1p3, 5, 0, STiles(2))
    Public Player10 As New Clan(C2p1, 0, 15, STiles(3))
    Public Player11 As New Clan(C2p2, 2, 15, STiles(4))
    Public Player12 As New Clan(C2p3, 4, 15, STiles(5))

    ' Creates a game and sets the board
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AssignCoord(0, 6)
        AssignCoord(0, 8)
        AssignCoord(1, 6)
        AssignCoord(1, 7)
        AssignCoord(1, 8)
        AssignCoord(2, 6)
        AssignCoord(2, 7)
        AssignCoord(2, 8)
        AssignCoord(3, 7)
        AssignCoord(3, 8)
        AssignCoord(4, 6)
        AssignCoord(4, 7)
        AssignCoord(4, 8)
        AssignCoord(5, 6)
        AssignCoord(5, 7)
        AssignCoord(5, 8)
        'RandomizeGems()
    End Sub

    ' Randomize gems on board
    Public Sub RandomizeGems()



        Dim LegalGemHexs As Hex() = {Board(0, 6), Board(0, 8),
                                     Board(1, 6), Board(1, 7), Board(1, 8),
                                     Board(2, 6), Board(2, 7), Board(2, 8),
                                                  Board(3, 7), Board(3, 8),
                                     Board(4, 6), Board(4, 7), Board(4, 8),
                                     Board(5, 6), Board(5, 7), Board(5, 8)}

        Dim LegalGemTiles As PictureBox() = {X0y6, X0y8,
                                             X1y6, X1y7, X1y8,
                                             X2y6, X2y7, X2y8}

        For i As Integer = 0 To 2 Step 1
            Dim rand As Integer = GetRandom(0, 7)
            Dim pb As New PictureBox()

            While (LegalGemHexs(rand).occupied = True)
                rand = GetRandom(0, 7)
            End While

            Dim hexagon As Hex = LegalGemHexs(rand)

            hexagon.gem = 5
            hexagon.occupied = True

            With pb
                .Tag = hexagon
                .Height = 65
                .Width = 60
                .SizeMode = PictureBoxSizeMode.StretchImage
                .BackColor = Color.Transparent
                .Image = My.Resources._5
                .Name = "Gem5_" & i
                .Location = New Point(CInt(LegalGemTiles(rand).Left + LegalGemTiles(rand).Width / 120),
                                              CInt(LegalGemTiles(rand).Top + LegalGemTiles(rand).Height / 120))
            End With

            LegalGemHexs(rand).occupied = True
            LegalGemHexs(rand).gem = 5
            AddHandler pb.Click, AddressOf Gem5_Click
            Controls.Add(pb)
            pb.BringToFront()
        Next

    End Sub

    ' Uses a static variable to allow random values based off of system clock 
    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Static Rand As System.Random = New System.Random()
        Return Rand.Next(Min, Max)
    End Function

    ' Updates a game when a move is made
    Public Sub MovePiece(Hexagon As Hex, NewCoord As Coord, Tile As PictureBox)
        If LegalMove(Board(NewCoord.x, NewCoord.y), NewCoord) Then
            Board(Current.GetCoord.x, Current.GetCoord.y).occupied = False
            Current.SetCoord = NewCoord
            Board(NewCoord.x, NewCoord.y).occupied = True
            Current.SetTile = Tile
            PicBox.Location = New Point(CInt(Tile.Left + Tile.Width / 180), CInt(Tile.Top + Tile.Height / 120))
            Moves -= 1
            lblMoves.Text = Moves
        End If
    End Sub

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
            lblScore.Text = CStr(Current.GetScore)
        End If

        Return legal
    End Function

    ' Sets the gamestate up for the next players turn
    Public Sub EndTurn()
        Moves = 0
        If PTurn = 0 Then
            PTurn = 1
        Else
            PTurn = 0
        End If
        Current = Nothing
        Dice_Thrown = False
        Player = If(Player = 1, 2, 1)
        lblPlayer.Text = "Player " & Player
    End Sub

    '''                      '''
    ''' PLAYER CLICK EVENTS  '''
    '''                      '''

    Private Sub C1p1_Click(sender As Object, e As EventArgs) Handles C1p1.Click
        Dim THex As Coord
        THex.x = Player00.GetCoord.x
        THex.y = Player00.GetCoord.y
        If PTurn = 0 Then
            Current = Player00
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player00.GetCoord.x, Player00.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player00.GetTile)
            C1p1.Image = Nothing
            C1p1.Dispose()
        End If
    End Sub

    Private Sub C1p2_Click(sender As Object, e As EventArgs) Handles C1p2.Click
        Dim THex As Coord
        THex.x = Player01.GetCoord.x
        THex.y = Player01.GetCoord.y
        If PTurn = 0 Then
            Current = Player01
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player01.GetCoord.x, Player01.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player01.GetTile)
            C1p2.Image = Nothing
            C1p2.Dispose()
        End If
    End Sub

    Private Sub C1p3_Click(sender As Object, e As EventArgs) Handles C1p3.Click
        Dim THex As Coord
        THex.x = Player02.GetCoord.x
        THex.y = Player02.GetCoord.y
        If PTurn = 0 Then
            Current = Player02
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player02.GetCoord.x, Player02.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player02.GetTile)
            C1p2.Image = Nothing
            C1p2.Dispose()
        End If
    End Sub

    Private Sub C2p1_Click(sender As Object, e As EventArgs) Handles C2p1.Click
        Dim THex As Coord
        THex.x = Player10.GetCoord.x
        THex.y = Player10.GetCoord.y
        If PTurn = 1 Then
            Current = Player10
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player10.GetCoord.x, Player10.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player11.GetTile)
            C2p1.Image = Nothing
            C2p1.Dispose()
        End If
    End Sub

    Private Sub C2p2_Click(sender As Object, e As EventArgs) Handles C2p2.Click
        Dim THex As Coord
        THex.x = Player11.GetCoord.x
        THex.y = Player11.GetCoord.y
        If PTurn = 1 Then
            Current = Player11
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player11.GetCoord.x, Player11.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player11.GetTile)
            C2p2.Image = Nothing
            C2p2.Dispose()
        End If
    End Sub

    Private Sub C2p3_Click(sender As Object, e As EventArgs) Handles C2p3.Click
        Dim THex As Coord
        THex.x = Player12.GetCoord.x
        THex.y = Player12.GetCoord.y
        If PTurn = 1 Then
            Current = Player12
            PicBox = DirectCast(sender, PictureBox)
        ElseIf LegalMove(Board(Player12.GetCoord.x, Player12.GetCoord.y), THex) Then
            MovePiece(Board(THex.x, THex.y), THex, Player12.GetTile)
            C2p3.Image = Nothing
            C2p3.Dispose()
        End If
    End Sub

    Private Sub Gem5_Click(sender As Object, e As EventArgs)
        Dim THex As Coord
        Dim hexagon As Hex = sender.Tag
        THex.x = hexagon.x
        THex.y = hexagon.y
        If LegalMove(hexagon, THex) Then
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Dice_Thrown = False Then
            Timer1.Start()
            Dice_Thrown = True
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar.Increment(5)
        If ProgressBar.Value = 100 Then
            Timer1.Stop()
            ProgressBar.Value = 0
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        EndTurn()
    End Sub
    ' remove this at the end
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Moves += 100
    End Sub


End Class
