Public Module Module1

    Public Structure Coord
        Public x As Integer
        Public y As Integer
    End Structure

    Public Structure Hex
        Public occupied As Boolean
        Public gem As Integer
        Public coordinate As Coord
        Public tile As PictureBox
        Public x As Integer
        Public y As Integer
    End Structure

    Public Random As Random
    Public Board(6, 16) As Hex
    Public Moves As Integer
    Public Dice_Thrown As Boolean = False
    Public Player = 1

End Module
