Attribute VB_Name = "modEditor"
Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const BLACK As Long = 0
Public Const WHITE As Long = 16777215


Public Tiles(100, 3) 'in the tileset. (includes regular and sprites)
'(index, property)
'properties:
'0 = type (1[tile] or 2[sprite])
'1 = tile number
'2 = tilename
'3 = tiledesc
Public totalTiles As Integer 'in the tileset. (includes regular AND sprites)
Public firstSpriteIndex As Integer 'first tile in tileset that is a sprite.
Public tileIndex(32767) As Integer 'this stores tile indexes when you know the tile number(2000 -> 6,etc.)


Public Sprites(100, 7) As Integer
'0 = location (row)
'1 = location (col)
'2 = tile number
'3-7 = attributes
Public numOfSprites As Integer
Public selSpriteIndex As Integer


Public Map(1000, 1000) As Integer
Public lvlWidth As Integer
Public lvlHeight As Integer

Public edge_LightColor As Long
Public edge_ShadowColor As Long

Public fill_Color As Long
Public pellet_Color As Long
Public fruit_Type As Integer



Public Function ClearLevelVars()

numOfSprites = 0

End Function


Public Function DeleteSelectedSprite()

    For i = selSpriteIndex To (numOfSprites - 1)
        For j = 0 To 7
            Sprites(i, j) = Sprites(i + 1, j)
        Next j
    Next i
    
    numOfSprites = numOfSprites - 1
    
End Function
