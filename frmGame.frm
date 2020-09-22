VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest for Honor"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5880
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6240
      Top             =   3240
   End
   Begin VB.TextBox txtSpeech 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmGame.frx":0000
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   3240
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label etqNToast 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   9
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label etqNTicket 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   8
         Top             =   2280
         Width           =   135
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   480
         X2              =   4200
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label etqNMagic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   7
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label etqNCoin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   6
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label etqNWood 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label etqTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.Image imgItem5 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":005A
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem4 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":0C9C
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStarGate 
      Height          =   480
      Left            =   720
      Picture         =   "frmGame.frx":18DE
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharSoldier 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2520
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem3 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":3162
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgItem2 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":3DA4
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryChar 
      Height          =   480
      Left            =   6120
      Picture         =   "frmGame.frx":49E6
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem1 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":5628
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen3 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":626A
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":6EAC
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed2 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":7AEE
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed1 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":8730
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable2 
      Height          =   480
      Left            =   4920
      Picture         =   "frmGame.frx":9372
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable1 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":9FB4
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB2 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":ABF6
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB1 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":B838
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgJar 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":C47A
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExit 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":D0BC
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":DCFE
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":E940
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":F582
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":101C4
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":10E06
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR3 
      Height          =   495
      Left            =   3000
      Picture         =   "frmGame.frx":11A48
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":126EA
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":1332C
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWall2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":13F6E
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeed 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":14BB0
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBadWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":157F2
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharGoodWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":16434
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy2 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":17076
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrees 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":17CB8
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRockHill 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":188FA
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":1953C
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":1A17E
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":1ADC0
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":1BA02
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFloor1 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":1C644
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallTop 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":1D286
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallBottom 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":1DEC8
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1EB0A
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1F74C
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":2038E
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":20FD0
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":21C12
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":22854
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":23496
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":240D8
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":24D1A
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":2595C
      Top             =   5640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNothing 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":2659E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   320
      Y2              =   0
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":271E0
      Top             =   5160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":27E22
      Top             =   5640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":28A64
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":296A6
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDownChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2A2E8
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2AF2A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeftChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":2BB6C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgUpChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":2C7AE
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":2D3F0
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":2E032
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":2EC74
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":2F8B6
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":304F8
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":3113A
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":31D7C
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":329BE
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":33600
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":34242
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":34E84
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":35AC6
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":36708
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":3734A
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":37F8C
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSign 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":38BCE
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'******************************************************
'Alex Whitney
'Quest for Honor
'
'Much thanks to Andres Zacarias for his ideas and a lot
'of artwork, as well as the map code (though the actual map
'is mine).
'Also I thank Simon Price for his DX 8 tutorial
'
'This code is not very well commented, so if you have
'any questions or comments, please e-mail me at:
'Milliardo@gundamwing.org
'
'Press escape to exit the game
'arrow keys to move
'i to access inventory
'space to talk
'and h for more help
'******************************************************
'******************************************************

Option Explicit



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If lost = True Then Unload Me
  Keyboard.GetDeviceStateKeyboard KeyboardState
  If KeyboardState.Key(DIK_UP) Then
    'Key Up.
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    CharFacing = 1
    Label1 = CharY
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY - 1
      DrawIt
      enctr
    End If
    If PositionMap = "#" Or PositionMap = "@" Then
      
      DrawIt
      Call enctr
    End If
  End If
  If KeyboardState.Key(DIK_DOWN) Then
    'Key Down.
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    CharFacing = 2
    Label1 = CharY
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY + 1
      DrawIt
      Call enctr
    End If
    If PositionMap = "M" Then
      
      DrawIt
    End If
End If
    If KeyboardState.Key(DIK_RIGHT) Then
    'Key Right.
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    CharFacing = 3
    Label2 = CharX
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX + 1
      DrawIt
      Call enctr
    End If
  End If
    If KeyboardState.Key(DIK_LEFT) Then
    'Key Left.
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    CharFacing = 4
    Label2 = CharX
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX - 1
      DrawIt
      Call enctr
    End If
  End If
  If KeyboardState.Key(DIK_SPACE) Then
    'Key Space.
    DrawIt
    
End If
If KeyboardState.Key(DIK_ESCAPE) Then
Unload Me
End If
If KeyboardState.Key(DIK_I) Then
    
    Load frminv
End If
If KeyboardState.Key(DIK_S) Then
    Call SaveGame
End If
If KeyboardState.Key(DIK_H) Then
    txtSpeech.Text = "S to save.  Arrow keys to move.  I to access inventory and stats.  R to rest to regain HP and MP."
    txtSpeech.Visible = True
End If
If KeyboardState.Key(DIK_R) Then
    mp = nmmp
    hp = nmhp
    fraMenu.Visible = True
    Timer2.Enabled = True
End If
End Sub


Private Sub Form_Load()
  frmGame.Height = 5175
  frmGame.Width = 4890
  Call InitGame
End Sub

Public Sub InitGame()
  Call itmdesc
  Set DI = DX.DirectInputCreate()
  Set Keyboard = DI.CreateDevice("GUID_SYSKEYBOARD")
  Set Mouse = DI.CreateDevice("GUID_SYSmouse")
  Keyboard.SetCommonDataFormat (DIFORMAT_KEYBOARD)
  Mouse.SetCommonDataFormat (DIFORMAT_MOUSE)
  Keyboard.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
  Mouse.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
  Keyboard.Acquire
  Mouse.Acquire
  'Character Position.
    MapLoaded = SAVE_MapLoaded
    SpeechLoaded = SAVE_SpeechLoaded
    CharX = SAVE_CharX
    CharY = SAVE_CharY
    CharFacing = SAVE_CharFacing
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    SpellCut = SAVE_SpellCut
    SpellDestroy = SAVE_SpellDestroy
  'Starting the first Map.
  If MapLoaded = "A1" Then Call Map_A1
  'Load the Speech for the Map.
  NewLine = Chr(13) + Chr(10)
 Call stats
hp = nmhp
mp = nmmp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set DI = Nothing
Set Keyboard = Nothing
Set Mouse = Nothing
Dim b As Integer
Set holding = Nothing
For b = 1 To 25
    Set itm(b) = Nothing
Next
End
End Sub



Private Sub Timer1_Timer()
  DrawIt
  Timer1.Enabled = False
End Sub

Public Sub DrawIt()
  'this code has to be credited to the talented Andres
  'Zacarias.  I Think it is quite good.
  For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler
        If X + CharX + 0 < 1 Then PictureHandler
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler
        If Y + CharY + 0 > 51 Then PictureHandler
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "b" '2 bush
        PaintPicture imgTrees.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "<" 'Weed
        PaintPicture imgWeed.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ">" 'Rock Hill
        PaintPicture imgRockHill.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "0" 'Cero Sign
        PaintPicture imgSign.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "q" 'grass
        PaintPicture imgTOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "A" 'Water
        PaintPicture imgBOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "a" 'grass
        PaintPicture imgBOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "w" 'grass
        PaintPicture imgTOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "S" 'Water
        PaintPicture imgBOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "s" 'grass
        PaintPicture imgBOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "e" 'Border Left grass
        PaintPicture ImgIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "r" 'Border Right grass
        PaintPicture ImgIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "d" 'Border Top grass
        PaintPicture ImgIT2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "F" 'Border Bottom water
        PaintPicture ImgIB.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "f" 'Border Bottom grass
        PaintPicture ImgIB2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "T" 'Border Bottom water
        PaintPicture ImgTIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "t" 'Border Bottom grass
        PaintPicture ImgTIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "Y" 'Border Bottom water
        PaintPicture ImgTIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "y" 'Border Bottom grass
        PaintPicture ImgTIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "V" 'Border Bottom water
        PaintPicture ImgBIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "v" 'Border Bottom grass
        PaintPicture ImgBIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "H" 'Border Bottom water
        PaintPicture ImgBIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "h" 'Border Bottom grass
        PaintPicture ImgBIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "U" 'Water
        PaintPicture ImgTIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "u" 'grass
        PaintPicture ImgIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "I" 'Water
        PaintPicture ImgTIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "i" 'grass
        PaintPicture ImgIB3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "J" 'Water
        PaintPicture ImgBIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "j" 'grass
        PaintPicture ImgIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "K" 'Water
        PaintPicture ImgBIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "k" 'grass
        PaintPicture ImgIT3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "," 'grass
        PaintPicture imgHouseExitB1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ";" 'grass
        PaintPicture imgHouseExitB2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "M" 'grass
        PaintPicture imgHouseExit.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "_" 'Nothing
        PaintPicture imgNothing.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Z" 'Wall bottom
        PaintPicture imgWallBottom.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "X" 'Wall top
        PaintPicture imgWallTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "C" 'Floor Blue 1
        PaintPicture imgFloor1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "z" 'Floor Tronco
        PaintPicture imgWall2.Picture, (X + 3) * 32, (Y + 3) * 32
            
      Case Is = "1" 'Laddy
        PaintPicture imgCharWomen1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "2" 'Boy 1
        PaintPicture imgCharBoy1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "3" 'Boy 2
        PaintPicture imgCharBoy2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "4" 'Good Wizard
        PaintPicture imgCharGoodWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "5" 'Bad Wizard
        PaintPicture imgCharBadWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "6" 'Laddy
        PaintPicture imgCharWomen2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "7" 'Laddy
        PaintPicture imgCharWomen3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "8" 'Soldier
        PaintPicture imgCharSoldier.Picture, (X + 3) * 32, (Y + 3) * 32
      
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "#" '3
        PaintPicture imgDoor1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "@" '3
        PaintPicture imgStarGate.Picture, (X + 3) * 32, (Y + 3) * 32
    
      Case Is = "!" 'Jar
        PaintPicture imgJar.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "(" 'Table
        PaintPicture imgTable1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ")" 'Table2
        PaintPicture imgTable2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "[" 'Bed
        PaintPicture imgBed1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "]" 'Bed2
        PaintPicture imgBed2.Picture, (X + 3) * 32, (Y + 3) * 32
    
    
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * 32, 5 * 32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * 32, 5 * 32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * 32, 5 * 32
  Case Is = 5
    PaintPicture imgCarryChar.Picture, 5 * 32, 5 * 32
  End Select
  
  Select Case Item
  Case Is = 1
    PaintPicture imgItem1.Picture, 5 * 32, 4 * 32
    Item = 0
  Case Is = 2
    PaintPicture imgItem2.Picture, 5 * 32, 4 * 32
    Item = 0
  Case Is = 3
    PaintPicture imgItem3.Picture, 5 * 32, 4 * 32
    Item = 0
  End Select

End Sub

Public Sub PictureHandler()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

Public Sub CharacterMovements()
  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * 32, 5 * 32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * 32, 5 * 32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * 32, 5 * 32
  End Select
End Sub

Private Sub Timer2_Timer()
fraMenu.Visible = False
MsgBox "HP full.  MP full.", vbOKOnly, "Rested!"
Timer2.Enabled = False
End Sub

