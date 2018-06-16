Public Class Clan
    Private score As Integer
    Private coordinate As Coord
    Private tile As String
    Private tag As Integer

    Public Sub New(ByVal Coord As Integer, ByVal Coord1 As Integer, ByVal Tile As String, ByVal Tag As Integer)
        Me.coordinate.x = Coord
        Me.coordinate.y = Coord1
        Me.tile = Tile
        Me.tag = Tag
    End Sub

    Public ReadOnly Property GetTag() As Integer
        Get
            Return Me.tag
        End Get
    End Property

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

    Public ReadOnly Property GetTile() As String
        Get
            Return Me.tile
        End Get
    End Property
End Class
