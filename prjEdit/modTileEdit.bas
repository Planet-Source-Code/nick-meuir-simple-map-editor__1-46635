Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public intTileRight As Integer
Public intTileBottom As Integer
Public Map(200, 200) As Byte
Public TileSelected As Integer
Public unloadfrm As Boolean

Sub SelectTile(intTileX As Integer, intTileY As Integer, intMode As Integer)
    If intMode = 1 Then
        intTileRight = intTileX * 32
        intTileBottom = intTileY * 32
    End If
End Sub
