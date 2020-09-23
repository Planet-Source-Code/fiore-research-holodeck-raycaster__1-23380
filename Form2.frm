VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Holodeck Startup"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Floor, Ceiling"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2520
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resolution"
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
      Begin VB.OptionButton Option5 
         Caption         =   "320 x 200"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "640 x 480"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
      Begin VB.CommandButton Command3 
         Caption         =   "Single Player"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Team DM"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CTF"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Deathmatch"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Multiplayer"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Level"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Line Line4 
         X1              =   1920
         X2              =   1920
         Y1              =   480
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   1920
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   120
         Y1              =   1320
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1920
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HELP: Arrow keys move, Q quits"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Floors = True
Else
    Floors = False
End If
End Sub

Private Sub Command1_Click()
    CD1.Filter = "Levels (*.dat)|*.dat|All Files (*.*)|*.*"
    CD1.DialogTitle = "Load Level"
    CD1.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNExplorer
    CD1.CancelError = True
    On Error Resume Next
    CD1.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    'Load the File
    LevelName = CD1.FileName
    LoadLevel (CD1.FileName)
    LevelLoaded = True
End Sub

Private Sub Command3_Click()
    If LevelLoaded = True Then
        PlayerX = (PlayerStartX * 64) - 32
        PlayerY = (PlayerStartY * 64) - 32
        ViewAngle = Angle90
        GameStarted = True
        Me.Hide
        Form1.Show
    Else
        Call MsgBox("You need to load a level first!", vbExclamation, "WARNING!!")
    End If
End Sub

Private Sub Form_Load()
    SinglePlayer = True
    GameStarted = False
    useTextures = True
    Floors = False
    Stride = 10
End Sub

Private Sub Option4_Click()
    RES = 640
End Sub

Private Sub Option5_Click()
    RES = 320
End Sub
Public Sub LoadLevel(FileName As String)
    Dim CTFVal, DMVal, TDMVal, SPVal
    Dim FileNum As Long
    FileNum = FreeFile
    Open FileName For Input As FileNum
    Do
        Dim Value As String
        Line Input #FileNum, Value
        Value = TrimExtraSpace(Value)
        If Trim$(Left$(Value, 1)) <> "" And Trim$(Left$(Value, 1)) <> "#" Then
            ReDim World(Val(Field(Value, "|", 1)), Val(Field(Value, "|", 2)))
            WorldSizeX = Val(Field(Value, "|", 1))
            WorldSizeY = Val(Field(Value, "|", 2))
            NumBots = Val(Field(Value, "|", 3))
            CTFVal = Val(Field(Value, "|", 4))
            DMVal = Val(Field(Value, "|", 5))
            TDMVal = Val(Field(Value, "|", 6))
            SPVal = Val(Field(Value, "|", 7))
            TexturePak = Field(Value, "|", 8)
            WeaponSet = Field(Value, "|", 9)
            PlayerStartX = Val(Field(Value, "|", 10))
            PlayerStartY = Val(Field(Value, "|", 11))
            ExitX = Val(Field(Value, "|", 12))
            ExitY = Val(Field(Value, "|", 13))
            BlueFlagX = Val(Field(Value, "|", 14))
            BlueFlagY = Val(Field(Value, "|", 15))
            RedFlagX = Val(Field(Value, "|", 16))
            RedFlagY = Val(Field(Value, "|", 17))
            CeilingTex = Val(Field(Value, "|", 18))
            FloorTex = Val(Field(Value, "|", 19))
            If CTFVal = 1 Then
                Option3.Enabled = True
            End If
            If DMVal = 1 Then
                Option1.Enabled = True
            End If
            If TDMVal = 1 Then
                Option2.Enabled = True
            End If
            If SPVal = 1 Then
                Option1.Enabled = False
                Option2.Enabled = False
                Option3.Enabled = False
                Command2.Enabled = False
                Command3.Enabled = False
                SinglePlayer = True
            End If
            Exit Do
        End If
    Loop
    For i = 1 To WorldSizeY
        Do
            Line Input #FileNum, Value
            Value = TrimExtraSpace(Value)
            If Trim$(Left$(Value, 1)) <> "" And Trim$(Left$(Value, 1)) <> "#" Then
                For ii = 1 To WorldSizeX
                    z = ii
                    World(ii, i) = Val(Field(Value, "|", z))
                Next ii
                Exit Do
            End If
        Loop
    Next i
    'Textures.Picture = LoadPicture(Dirpath & TexturePak)
    'PicGun1.Picture = LoadPicture(Dirpath & WeaponSet & "1.bmp")
    'PicGun2.Picture = LoadPicture(Dirpath & WeaponSet & "2.bmp")
    'PicGun3.Picture = LoadPicture(Dirpath & WeaponSet & "3.bmp")
    'PicGun4.Picture = LoadPicture(Dirpath & WeaponSet & "4.bmp")
    'Continue work here Program Bot(AI)
    Close FileNum
End Sub

Public Function TrimExtraSpace(ByVal Valu As String) As String
    ' Remove leading and trailing spaces
    Dim flag As Boolean
    flag = False
    Valu = RTrim(LTrim(Valu))
    ' Loop through string and remove extra spaces between
    ' characters...
    For t& = 1 To Len(Valu)
        this$ = Mid$(Valu, t&, 1)
        If this$ <> " " Then
            If flag = True Then
                TrimExtraSpace = TrimExtraSpace + " "
                flag = False
            End If
            TrimExtraSpace = TrimExtraSpace + this$
        Else
            flag = True
        End If
    Next
End Function

Public Function Field(strg As String, sep As String, cnt) As String
    ' Finds occurance (as passed in "cnt" argument) inside of string "strg", in
    ' of fields delimited by "sep".
    ' Example:
    '   a$ = Field("XYZ|ZTX|ABC","|",2)
    '   returns
    '   a$ = "ZTX"
    Field = ""
    strg2$ = sep & strg & sep
    If cnt < 1 Or Len(sep) <> 1 Then Exit Function
    ls& = Len(strg2$)
    For t& = 1 To ls&
        If Mid$(strg2$, t&, 1) = sep Then
            If cnt = 1 Then
                BegStrg& = t& + 1
            ElseIf cnt = 0 Then
                EndStrg& = t& - 1
                Field = Mid$(strg2$, BegStrg&, EndStrg& - BegStrg& + 1)
                Exit For
            End If
            cnt = cnt - 1
        End If
    Next
End Function

