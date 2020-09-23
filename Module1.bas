Attribute VB_Name = "Module1"
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Const CellY = 64
Public Const CellX = 64
Public Const ProjectionWidth = 320
Public Const ProjectionHeight = 200
Public Const Angle60 = ProjectionWidth
Public Const Angle30 = Angle60 / 2
Public Const Angle15 = Angle30 / 2
Public Const Angle90 = Angle30 * 3
Public Const Angle180 = Angle90 * 2
Public Const Angle270 = Angle90 * 3
Public Const Angle360 = Angle60 * 6
Public Const Angle0 = 0
Public Const Angle6 = Angle30 / 5
Public Const Angle5 = Angle30 / 6
Public Const Angle10 = Angle5 * 2
Public Const InvCellX = 1 / CellX
Public Const InvCellY = 1 / CellY
Public Const MinDistance = 28
Public Const FocalDist = 277

Public RES
Public World() As Single
Public Stride
Public TanTable(1920)
Public CoTanTable(1920)
Public PixelYTable(200)
Public TextureTable(10)
Public CosTable(1920)
Public InvCosTable(1920)
Public SinTable(1920)
Public WallHeightTable(100000)
Public PlayerX, PlayerY, ViewAngle
Public FPS
Public PlayerStartX, PlayerStartY, TextureNumber, TextureColumn
Public Ray, Skale, UpperEnd, LowerEnd
Public Rightb As Boolean
Public Leftb As Boolean
Public Upb As Boolean
Public Downb As Boolean
Public Fireb As Boolean
Public DM As Boolean
Public TDM As Boolean
Public CTF As Boolean
Public ExitX, ExitY
Public BlueFlagX, BlueFlagY, RedFlagX, RedFlagY
Public CeilingTex, FloorTex
Public SinglePlayer As Boolean
Public LevelLoaded As Boolean
Public WorldSizeY, WorldSizeX
Public useTextures As Boolean
Public EnemyX, EnemyY
Public GameStarted As Boolean
Public LevelName As String
Public TexturePak, WeaponSet
Public Floors As Boolean
