VERSION 5.00
Begin VB.Form frminv 
   BackColor       =   &H0000C000&
   Caption         =   "Inventory"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   Picture         =   "frminv.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Pts 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   39
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ptl 
      Caption         =   "Points Left:"
      Height          =   255
      Left            =   4440
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label ld 
      Caption         =   "Defensive"
      Height          =   255
      Left            =   5280
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lo 
      Caption         =   "Offensive"
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lsp 
      Caption         =   "Learn new spell"
      Height          =   255
      Left            =   5040
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image pdex 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4200
      Picture         =   "frminv.frx":0C42
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pos 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6120
      Picture         =   "frminv.frx":0F84
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pod 
      Enabled         =   0   'False
      Height          =   240
      Left            =   6120
      Picture         =   "frminv.frx":12C6
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pcon 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4200
      Picture         =   "frminv.frx":1608
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pint 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4200
      Picture         =   "frminv.frx":194A
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pstr 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4200
      Picture         =   "frminv.frx":1C8C
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lbdmg 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   34
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "DMG"
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   3360
      TabIndex        =   32
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image blnk 
      Height          =   495
      Left            =   5280
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label llvl 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   960
      Width           =   255
   End
   Begin VB.Label llvld 
      Caption         =   "Level"
      Height          =   255
      Left            =   1320
      TabIndex        =   29
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lac 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lacd 
      Caption         =   "AC"
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   2160
      Width           =   255
   End
   Begin VB.Image Ihlm3 
      Height          =   480
      Left            =   6240
      Picture         =   "frminv.frx":1FCE
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ihlm2 
      Height          =   480
      Left            =   5760
      Picture         =   "frminv.frx":2C10
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ihlm1 
      Height          =   480
      Left            =   5280
      Picture         =   "frminv.frx":3852
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg6 
      Height          =   480
      Left            =   3960
      Picture         =   "frminv.frx":4494
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg5 
      Height          =   480
      Left            =   3480
      Picture         =   "frminv.frx":50D6
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg4 
      Height          =   480
      Left            =   3000
      Picture         =   "frminv.frx":5D18
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg3 
      Height          =   480
      Left            =   2520
      Picture         =   "frminv.frx":695A
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg2 
      Height          =   480
      Left            =   2040
      Picture         =   "frminv.frx":759C
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ileg1 
      Height          =   480
      Left            =   1560
      Picture         =   "frminv.frx":81DE
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iear1 
      Height          =   480
      Left            =   5760
      Picture         =   "frminv.frx":8E20
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Irng1 
      Height          =   480
      Left            =   5280
      Picture         =   "frminv.frx":9A62
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ishd4 
      Height          =   480
      Left            =   3000
      Picture         =   "frminv.frx":A6A4
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ishd3 
      Height          =   480
      Left            =   2520
      Picture         =   "frminv.frx":B2E6
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ishd2 
      Height          =   480
      Left            =   2040
      Picture         =   "frminv.frx":BF28
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Ishd1 
      Height          =   480
      Left            =   1560
      Picture         =   "frminv.frx":CB6A
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iwpn4 
      Height          =   480
      Left            =   3000
      Picture         =   "frminv.frx":D7AC
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iwpn3 
      Height          =   480
      Left            =   2520
      Picture         =   "frminv.frx":E3EE
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iwpn2 
      Height          =   480
      Left            =   2040
      Picture         =   "frminv.frx":F030
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iwpn1 
      Height          =   480
      Left            =   1560
      Picture         =   "frminv.frx":FC72
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iarm6 
      Height          =   480
      Left            =   3960
      Picture         =   "frminv.frx":108B4
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iarm5 
      Height          =   480
      Left            =   3480
      Picture         =   "frminv.frx":114F6
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iarm4 
      Height          =   480
      Left            =   3000
      Picture         =   "frminv.frx":12138
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iarm3 
      Height          =   480
      Left            =   2520
      Picture         =   "frminv.frx":12D7A
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Iarm2 
      Height          =   480
      Left            =   2040
      Picture         =   "frminv.frx":139BC
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Ldesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   3360
      TabIndex        =   26
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Iarm1 
      Height          =   480
      Left            =   1560
      Picture         =   "frminv.frx":145FE
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   2640
      Y2              =   4080
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   1680
      Y1              =   5040
      Y2              =   4080
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   1680
      Y1              =   5040
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   2520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   1680
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   1920
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line7 
      X1              =   2160
      X2              =   1920
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line8 
      X1              =   2280
      X2              =   2160
      Y1              =   1920
      Y2              =   2280
   End
   Begin VB.Line Line9 
      X1              =   2160
      X2              =   2280
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line10 
      X1              =   1920
      X2              =   2160
      Y1              =   1440
      Y2              =   1560
   End
   Begin VB.Line Line11 
      X1              =   1560
      X2              =   1920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line12 
      X1              =   1200
      X2              =   1560
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line13 
      X1              =   1080
      X2              =   1200
      Y1              =   1800
      Y2              =   1560
   End
   Begin VB.Line Line14 
      X1              =   1200
      X2              =   1080
      Y1              =   2160
      Y2              =   1800
   End
   Begin VB.Line Line15 
      X1              =   1440
      X2              =   1200
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line Line16 
      X1              =   1680
      X2              =   1440
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   23
      Left            =   1440
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   18
      Left            =   360
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   20
      Left            =   2520
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   19
      Left            =   1440
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   15
      Left            =   1440
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   22
      Left            =   2520
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   21
      Left            =   360
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   17
      Left            =   2280
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image inv 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   16
      Left            =   600
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Earring"
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Helmet"
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Earring"
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Armor"
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Shield"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ring"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Leggings"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Ring"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label100 
      Caption         =   "STR"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lstr 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label ldex 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label ldexd 
      Caption         =   "DEX"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lcond 
      Caption         =   "CON"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lcon 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lintd 
      Caption         =   "INT"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lint 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lexpd 
      Caption         =   "EXP"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lexp 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lhpd 
      Caption         =   "HP"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lmpd 
      Caption         =   "MP"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lhp 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lmhp 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lmp 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lmmp 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   375
   End
   Begin VB.Line Line17 
      X1              =   4080
      X2              =   3960
      Y1              =   4080
      Y2              =   4320
   End
   Begin VB.Line Line18 
      X1              =   4080
      X2              =   3960
      Y1              =   3840
      Y2              =   4080
   End
End
Attribute VB_Name = "frminv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xt As Boolean


Private Sub Command1_Click()
xt = True
End Sub



Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keyboard.GetDeviceStateKeyboard keyboard2state
    If keyboard2state.Key(DIK_I) Then xt = True

End Sub

Private Sub Form_Load()
frmGame.Hide
frminv.Height = 5805
frminv.Width = 6750
frminv.Top = 0
frminv.Left = 0
picinit
Dim a As Integer
inv(19) = itm(m)
inv(18) = itm(w)
inv(20) = itm(s)
inv(21) = itm(r1)
inv(22) = itm(r2)
inv(16) = itm(e1)
inv(17) = itm(e2)
inv(23) = itm(l)
inv(15) = itm(h)
Show
mainloop
unloadinv
End Sub
Sub picinit()
Set itm(0) = blnk.Picture
Set itm(1) = Iarm1.Picture
Set itm(2) = Iarm2.Picture
Set itm(3) = Iarm3.Picture
Set itm(4) = Iarm4.Picture
Set itm(5) = Iarm5.Picture
Set itm(6) = Iarm6.Picture
Set itm(7) = Iwpn1.Picture
Set itm(8) = Iwpn2.Picture
Set itm(9) = Iwpn3.Picture
Set itm(10) = Iwpn4.Picture
Set itm(11) = Ishd1.Picture
Set itm(12) = Ishd2.Picture
Set itm(13) = Ishd3.Picture
Set itm(14) = Ishd4.Picture
Set itm(15) = Irng1.Picture
Set itm(16) = Iear1.Picture
Set itm(17) = Ileg1.Picture
Set itm(18) = Ileg2.Picture
Set itm(19) = Ileg3.Picture
Set itm(20) = Ileg4.Picture
Set itm(21) = Ileg5.Picture
Set itm(22) = Ileg6.Picture
Set itm(23) = Ihlm1.Picture
Set itm(24) = Ihlm2.Picture
Set itm(25) = Ihlm3.Picture
End Sub
Sub mainloop()
Do
    DoEvents
    If hp > nmhp Then hp = nmhp
    If mp > nmmp Then mp = nmmp
    CheckLvl
    lstr.Caption = str
    ldex.Caption = dex
    lcon.Caption = con
    lint.Caption = inte
    lexp.Caption = EXPer
    lhp.Caption = hp
    lmp.Caption = mp
    lmhp.Caption = nmhp
    lmmp.Caption = nmmp
    lac.Caption = nac
    llvl.Caption = level
    lbdmg.Caption = wdmg & "+" & bst
    Pts.Caption = lpt
    If xt = True Then xt = False: Exit Do
Loop
End Sub

Sub unloadinv()
frmGame.Show
Unload frminv
Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Visible = False
Ldesc.Caption = ""
End Sub

Private Sub inv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case inv(Index).Picture 'if the mouse moves over an
                               'inv slot of eq slot then check
                               'to see if it has a picture in it
    Case Is = Iarm1.Picture 'if the picture in it is this
        Ldesc.Caption = darm1 'desc is _Predetermined_
        Ldesc.Visible = True 'make ldesc visible
    Case Is = Iarm2.Picture
        Ldesc.Caption = darm2
        Ldesc.Visible = True
    Case Is = Iarm3.Picture
        Ldesc.Caption = darm3
        Ldesc.Visible = True
    Case Is = Iarm4.Picture
        Ldesc.Caption = darm4
        Ldesc.Visible = True
    Case Is = Iarm5.Picture
        Ldesc.Caption = darm5
        Ldesc.Visible = True
    Case Is = Iarm6.Picture
        Ldesc.Caption = darm6
        Ldesc.Visible = tr
    Case Is = Ihlm1.Picture
        Ldesc.Caption = dhlm1
        Ldesc.Visible = True
    Case Is = Ihlm2.Picture
        Ldesc.Caption = dhlm2
        Ldesc.Visible = True
    Case Is = Ihlm3.Picture
        Ldesc.Caption = dhlm3
        Ldesc.Visible = True
    Case Is = Irng1.Picture
        Ldesc.Caption = drng1
        Ldesc.Visible = True
    Case Is = Iear1.Picture
        Ldesc.Caption = dear1
        Ldesc.Visible = True
    Case Is = Ishd1.Picture
        Ldesc.Caption = dshd1
        Ldesc.Visible = True
    Case Is = Ishd2.Picture
        Ldesc.Caption = dshd2
        Ldesc.Visible = True
    Case Is = Ishd3.Picture
        Ldesc.Caption = dshd3
        Ldesc.Visible = True
    Case Is = Ishd4.Picture
        Ldesc.Caption = dshd4
        Ldesc.Visible = True
    Case Is = Iwpn1.Picture
        Ldesc.Caption = dwpn1
        Ldesc.Visible = True
    Case Is = Iwpn2.Picture
        Ldesc.Caption = dwpn2
        Ldesc.Visible = True
    Case Is = Iwpn3.Picture
        Ldesc.Caption = dwpn3
        Ldesc.Visible = True
    Case Is = Iwpn4.Picture
        Ldesc.Caption = dwpn4
        Ldesc.Visible = True
    Case Is = Ileg1.Picture
        Ldesc.Caption = dleg1
        Ldesc.Visible = True
    Case Is = Ileg2.Picture
        Ldesc.Caption = dleg2
        Ldesc.Visible = True
    Case Is = Ileg3.Picture
        Ldesc.Caption = dleg3
        Ldesc.Visible = True
    Case Is = Ileg4.Picture
        Ldesc.Caption = dleg4
        Ldesc.Visible = True
    Case Is = Ileg5.Picture
        Ldesc.Caption = dleg5
        Ldesc.Visible = True
    Case Is = Ileg6.Picture
        Ldesc.Caption = dleg6
        Ldesc.Visible = True
    Case Is = 0
    Case Else
        Ldesc.Caption = "ERROR!  Description not found."
        Ldesc.Visible = True
End Select
    
    
    
    
Ldesc.Visible = True
End Sub

Private Sub Label100_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This stat affects your damage."
Ldesc.Visible = True
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "The number before the + is the dice roll.  The number after the + is what is added to the roll."
Ldesc.Visible = True
End Sub

Private Sub lacd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This is your total amount of Armor Class.  This determines how well you are defended against enemy attacks.  The higher the better."
Ldesc.Visible = True
End Sub

Private Sub lcond_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This stat affects your Hit Points."
Ldesc.Visible = True
End Sub


Private Sub ldexd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This stat gives you bonuses to your Armor Class, as well as determining how well you hit opponents."
Ldesc.Visible = True

End Sub

Private Sub lexpd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This is how much Experience you have.  The more you have, the higher level you will be, and the stronger you will get."
Ldesc.Visible = True
End Sub

Private Sub lhpd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "The number on the left is the amount of Hit Points you have now.  The number on the right is the maximum amount you can have.  Once you reach 0, you are dead."
Ldesc.Visible = True
End Sub

Private Sub lintd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This affects how much mana you have.  The more mana you have, the more spells you can cast."
Ldesc.Visible = True
End Sub

Private Sub llvld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "This is your level.  The higher level you are, the stronger you have become."
Ldesc.Visible = True
End Sub

Private Sub lmpd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ldesc.Caption = "The number on the left is your current Mana Points.  The number on the right is your maximum number of mana points.  It takes Mana to cast spells.  When you are out, you cannot cast anymore spells."
Ldesc.Visible = True
End Sub

Private Sub pcon_Click()
con = con + 1
lpt = lpt - 1
Call stats
End Sub

Private Sub pdex_Click()
dex = dex + 1
lpt = lpt - 1
Call stats
End Sub

Private Sub pint_Click()
inte = inte + 1
lpt = lpt - 1
Call stats
End Sub

Private Sub pod_Click()
Dim k As Integer
For k = 0 To 3
    If dsp(k) = False Then
        dsp(k) = True
        lpt = lpt - 1
        Exit Sub
    End If
Next
End Sub

Private Sub pos_Click()
Dim k As Integer
For k = 0 To 3
    If osp(k) = False Then
        osp(k) = True
        lpt = lpt - 1
        Exit Sub
    End If
Next
End Sub

Private Sub pstr_Click()
str = str + 1
lpt = lpt - 1
Call stats
End Sub
Sub CheckLvl()
If lpt < 1 Then
    pod.Visible = False 'should have made objects arrays-
    pod.Enabled = False 'would have been MUCH easier here
    pos.Visible = False 'but i would need to redo code for
    pos.Enabled = False 'the clicking!
    pstr.Visible = False
    pstr.Enabled = False
    pdex.Visible = False
    pdex.Enabled = False
    pcon.Visible = False
    pcon.Enabled = False
    pint.Visible = False
    pint.Enabled = False
    lsp.Visible = False
    lo.Visible = False
    ld.Visible = False
    Pts.Visible = False
    ptl.Visible = False
End If
If lpt >= 1 Then
    pod.Visible = True
    pod.Enabled = True
    pos.Visible = True
    pos.Enabled = True
    pstr.Visible = True
    pstr.Enabled = True
    pdex.Visible = True
    pdex.Enabled = True
    pcon.Visible = True
    pcon.Enabled = True
    pint.Visible = True
    pint.Enabled = True
    lsp.Visible = True
    lo.Visible = True
    ld.Visible = True
    Pts.Visible = True
    ptl.Visible = True
End If

    
End Sub
