Public Class Clan
    Private picture As PictureBox
    Private score As Integer
    Private coordinate As Coord
    Private tile As PictureBox

    Public Sub New(ByVal obj As Object, ByVal Coord As Integer, ByVal Coord1 As Integer, ByVal Tile As PictureBox)
        Me.picture = obj
        Me.coordinate.x = Coord
        Me.coordinate.y = Coord1
        Me.tile = Tile
    End Sub

    Public WriteOnly Property SetScore() As Integer
        Set(ByVal Value As Integer)
            Me.score += Value
        End Set
    End Property

    Public ReadOnly Property GetScore As Integer
        Get
            Return Me.score
        End Get
    End Property

    Public ReadOnly Property GetCoord() As Coord
        Get
            Return Me.coordinate
        End Get
    End Property

    Public WriteOnly Property SetCoord() As Coord
        Set(ByVal Value As Coord)
            Me.coordinate.x = Value.x
            Me.coordinate.y = Value.y
        End Set
    End Property

    Public ReadOnly Property GetPic() As PictureBox
        Get
            Return Me.picture
        End Get
    End Property

    Public WriteOnly Property SetTile() As PictureBox
        Set(ByVal Value As PictureBox)
            Me.tile = Value
        End Set
    End Property

    Public ReadOnly Property GetTile() As PictureBox
        Get
            Return Me.tile
        End Get
    End Property
End Class
