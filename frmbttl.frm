VERSION 5.00
Begin VB.Form frmbttl 
   BackColor       =   &H0000C000&
   Caption         =   "Battle!"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer dly 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   5880
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   5520
      Top             =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Shielding"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Lshd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Ispl 
      Height          =   975
      Left            =   1800
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label dspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armor"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image kill 
      Height          =   960
      Index           =   1
      Left            =   6480
      Picture         =   "frmbttl.frx":0000
      Top             =   5640
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label intr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prepare to Fight!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   20
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Image Fshield 
      Height          =   960
      Left            =   1080
      Picture         =   "frmbttl.frx":3042
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image mshield 
      Height          =   960
      Left            =   2640
      Picture         =   "frmbttl.frx":5B84
      Top             =   4680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image HEal 
      Height          =   1035
      Left            =   840
      Picture         =   "frmbttl.frx":7EC6
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image shield 
      Height          =   960
      Left            =   1920
      Picture         =   "frmbttl.frx":AD64
      Top             =   4680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label dspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Flame Shield"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label dspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mana Shield"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label dspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Heal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Cancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label despl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Defensive"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Ofspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Offensive"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label bttlinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image Earthrift 
      Height          =   1020
      Left            =   4440
      Picture         =   "frmbttl.frx":D0A6
      Top             =   2880
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Image iceblast 
      Height          =   1020
      Left            =   4440
      Picture         =   "frmbttl.frx":19248
      Top             =   4080
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Image lgtstrike 
      Height          =   1110
      Left            =   4440
      Picture         =   "frmbttl.frx":253EA
      Top             =   4320
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Image Flmspin 
      Height          =   1005
      Left            =   4440
      Picture         =   "frmbttl.frx":337F4
      Top             =   3240
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.Label ospl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Earth Rift"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label ospl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ice Blast"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label ospl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lightning Strike"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line ln 
      Index           =   1
      Visible         =   0   'False
      X1              =   4320
      X2              =   2520
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line ln 
      Index           =   2
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   2880
      Y2              =   4440
   End
   Begin VB.Line ln 
      Index           =   0
      Visible         =   0   'False
      X1              =   4320
      X2              =   4320
      Y1              =   4440
      Y2              =   2880
   End
   Begin VB.Line ln 
      Index           =   3
      Visible         =   0   'False
      X1              =   4320
      X2              =   2520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label ospl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Flame Spin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label ehp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label php 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label pmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Run 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label spl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spell"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label atk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Mon 
      Height          =   960
      Index           =   0
      Left            =   5160
      Picture         =   "frmbttl.frx":41A56
      Top             =   1680
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6480
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   1200
      Picture         =   "frmbttl.frx":44A98
      Top             =   2160
      Width           =   480
   End
End
Attribute VB_Name = "frmbttl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v As Integer
Dim rn As Boolean
Dim bti As String
Dim dnct As Integer
Dim dn As Boolean
Dim seff As Integer
Dim Fshielda As Boolean

Private Sub atk_Click()
swing
dly.Enabled = True
End Sub

Private Sub Cancel_Click()
If Ofspl.Visible = True Then
    atk.Visible = True
    Run.Visible = True
    spl.Visible = True
    Run.Enabled = True
    spl.Enabled = True
    atk.Enabled = True
    Ofspl.Visible = False
    Ofspl.Enabled = False
    despl(0).Enabled = False
    despl(0).Visible = False
    Cancel.Visible = False
    Cancel.Enabled = False
End If
If ln(0).Visible = True Then
    Dim l As Integer
    For l = 0 To 3
        ln(l).Visible = False
        dspl(l).Visible = False
        dspl(l).Enabled = False
        ospl(l).Visible = False
        ospl(l).Enabled = False
    Next
    Ofspl.Visible = True
    Ofspl.Enabled = True
    despl(0).Enabled = True
    despl(0).Visible = True
End If
End Sub

Private Sub despl_Click(Index As Integer)
Dim l As Integer
For l = 0 To 3
    ln(l).Visible = True
    If dsp(l) = True Then
        dspl(l).Visible = True
        dspl(l).Enabled = True
    End If
Next
despl(0).Visible = False
despl(0).Enabled = False
Ofspl.Enabled = False
Ofspl.Visible = False
End Sub

Private Sub dly_Timer()
If dn = True Then
    dn = False
    rn = False
    Unload Me
Else
    cswing
    dly.Enabled = False
End If
Ispl.Visible = False
HEal.Visible = False
End Sub
Private Sub dspl_Click(Index As Integer)
Randomize
Dim ddesc As String
Dim pHEal As Integer
Select Case Index
    Case 0
        If mp - 2 >= 0 Then
                mp = mp - 2
                seff = (Rnd * 4) + 1 + bsd + seff
                Ispl.Picture = shield.Picture
                Ispl.Visible = True
                ddesc = "You feel shielded (" & seff & ")"
        Else: Exit Sub
        End If
    Case 1
        If mp - 3 >= 0 Then
            mp = mp - 3
            pHEal = (Rnd * 6) + 1 + bsd
            hp = hp + pHEal
            HEal.Visible = True
            ddesc = "You have been healed " & pHEal & " points."
            Else: Exit Sub
        End If
    Case 2
        If mp - 5 >= 0 Then
                mp = mp - 5
                seff = mp + seff
                Ispl.Picture = mshield.Picture
                Ispl.Visible = True
                ddesc = "You feel shielded (" & mp & ")"
            Else: Exit Sub
        End If
    Case 3
        If mp - 7 >= 0 Then
                mp = mp - 7
                seff = (Rnd * 12) + 1 + bsd + seff
                Ispl.Picture = Fshield.Picture
                Ispl.Visible = True
                ddesc = "You feel a warm shielding (" & seff & ")"
                Fshield = True
            Else: Exit Sub
        End If
End Select
bttlinfo.Caption = ddesc
bttlinfo.Visible = True
dly.Enabled = True
Cancel_Click
Cancel_Click
End Sub

Private Sub Form_Load()
seff = 0
dcnt = 0
frmbttl.Show
frmGame.Hide
Call stats
Call Mons
v = monst
frmbttl.Height = 5100
frmbttl.Width = 6930
frmbttl.Top = 0
frmbttl.Left = 0
Timer1.Enabled = True
mainloop
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Randomize
Dim itmmsg As String
Dim gitm As Integer
gitm = Int(Rnd * 9) + 1
Select Case gitm
    Case Is = 1
        If m <= 5 Then
        m = m + 1
        itmmsg = "You have acquired new body armor!"
        End If
    Case Is = 2
        If w <= 2 Then
        w = w + 1
        itmmsg = "You have acquired a new weapon!"
        End If
    Case Is = 3
        If s <= 3 Then
        s = s + 1
        itmmsg = "You have acquired a new shield!"
        End If
    Case Is = 5
        If r2 = 0 Then
        r2 = 15
        itmmsg = "You have acquired a new ring!"
        End If
    Case Is = 7
        If e2 = 0 Then
        e2 = 16
        itmmsg = "You have acquired a new earring!"
        End If
    Case Is = 8
        If l <= 5 Then
        l = l + 1
        itmmsg = "You have acquired a new earring!"
        End If
    Case Is = 9
        If h <= 2 Then
        h = h + 1
        itmmsg = "You have acquired a new helmet!"
        End If
End Select
If itmmsg <> "" Then
    MsgBox itmmsg, vbOKOnly, "New item"
End If
itmmsg = ""
skp:
frmGame.Show
bttlinfo.Caption = ""
bttlinfo.Visible = False
Call stats
End Sub




Private Sub Ofspl_Click()
Dim l As Integer
For l = 0 To 3
    ln(l).Visible = True
    If osp(l) = True Then
        ospl(l).Visible = True
        ospl(l).Enabled = True
    End If
Next
Ofspl.Visible = False
Ofspl.Enabled = False
despl(0).Visible = False
despl(0).Enabled = False
End Sub

Private Sub ospl_Click(Index As Integer)
Randomize
Dim sdmg As Integer
Select Case Index
    Case 0
        If mp - 2 >= 0 Then
            mp = mp - 2
            sdmg = (Rnd * 2) + 1 + bsd
            Ispl.Picture = Flmspin.Picture
            Ispl.Visible = True
            Else: Exit Sub
        
        End If
    Case 1
        If mp - 4 >= 0 Then
            mp = mp - 4
            sdmg = (Rnd * 3) + 2 + bsd
            Ispl.Picture = lgtstrike.Picture
            Ispl.Visible = True
            Else: Exit Sub
        
        End If
    Case 2
        If mp - 5 >= 0 Then
            mp = mp - 5
            sdmg = (Rnd * 4) + 3 + bsd
            Ispl.Picture = iceblast.Picture
            Ispl.Visible = True
                    Else: Exit Sub

        End If
    Case 3
        If mp - 6 >= 0 Then
            mp = mp - 6
            sdmg = (Rnd * 5) + 4 + bsd
            Ispl.Picture = Earthrift.Picture
            Ispl.Visible = True
        Else: Exit Sub
        End If
End Select
bttlinfo.Visible = True
bttlinfo.Caption = "Your spell did " & sdmg & " damage!"
chp(v) = chp(v) - sdmg
 If chp(v) <= 0 Then
        EXPer = EXPer + ccr(v)
        MsgBox ("You have gained " & ccr(v) & " experience!"), vbInformation, "Victorious"
        dn = True
        kill(1).Visible = True
        spl.Visible = False
        spl.Enabled = False
        Run.Visible = False
        Run.Enabled = False
        atk.Enabled = False
        atk.Visible = False
        Exit Sub
    End If
Cancel_Click
Cancel_Click
dly.Enabled = True
End Sub



Private Sub Run_Click()
Dim trun
Randomize
trun = Int(Rnd * 6) + 1
If trun >= Int(ccr(v) / level) Then
    bttlinfo.Caption = "You have escaped!"
    bttlinfo.Visible = True
    dn = True
Else
    bttlinfo.Caption = "You have failed to escape!"
    bttlinfo.Visible = True
End If
dly.Enabled = True
End Sub

Private Sub spl_Click()
Ofspl.Visible = True
Ofspl.Enabled = True
despl(0).Enabled = True
despl(0).Visible = True
Cancel.Visible = True
Cancel.Enabled = True
spl.Visible = False
spl.Enabled = False
Run.Visible = False
Run.Enabled = False
atk.Enabled = False
atk.Visible = False
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
intr.Visible = False
Mon(v).Visible = True
Run.Visible = True
atk.Visible = True
spl.Visible = True
Run.Enabled = True
spl.Enabled = True
atk.Enabled = True
End Sub
Sub mainloop()
Do
    
    DoEvents
    If seff < 0 Then seff = 0
    Lshd.Caption = seff
    If rn = True Or dn = True Then
        rn = False
        dn = False
        Exit Sub
    End If
    Keyboard.GetDeviceStateKeyboard keyboard2state
    If keyboard2state.Key(DIK_Q) Then
        rn = True
    End If
    ehp.Caption = chp(v)
    php.Caption = hp
    pmp.Caption = mp
    If hp <= 0 Then
        MsgBox "You have been defeated!  Game over.", vbOKOnly, "Game Over"
        lost = True
        Call dsv
        Exit Do
    End If
    If seff = 0 Then
        Fshielda = False
    End If
Loop
End Sub
Sub atck()
Dim dm As Integer

Randomize
dm = Int(Rnd * wdmg) + 1
dam = dm + bst
chp(v) = chp(v) - dam
bti = "you hit for " & dam & " damage!"
bttlinfo.Caption = bti
bttlinfo.Visible = True
    If chp(v) <= 0 Then
        EXPer = EXPer + ccr(v)
        MsgBox ("You have gained " & ccr(v) & " experience!"), vbInformation, "Victorious"
        dn = True
        kill(1).Visible = True
        spl.Visible = False
        spl.Enabled = False
        Run.Visible = False
        Run.Enabled = False
        atk.Enabled = False
        atk.Visible = False
    End If

End Sub
Sub catck()
Dim cdmg As Integer
Dim cdam As Integer
Randomize
cdmg = Int(Rnd * cdm(v)) + 1
cdam = cdmg + cbs
If seff >= 1 Then
    Dim tseff As Integer
    tseff = seff
    tseff = seff - cdam
    cdam = cdam - seff
    seff = tseff
    tseff = 0
    If cdam < 0 Then cdam = 0
End If
If Fshielda = True Then
    chp(v) = chp(v) - 3
End If
hp = hp - cdam
bti = "Your opponent hit you for " & cdam & " damage!"
If cdam <= 0 Then
    bti = "Your shield blocked your opponents damage!"
End If
bttlinfo.Caption = bti
bttlinfo.Visible = True
End Sub
Sub swing()
Dim sw As Integer
Dim ncac As Integer
Randomize
sw = Int(Rnd * 20) + 1
ncac = cac(v) + 10
If (Int(dex / 2) + sw) >= ncac Then
    atck
Else
    bttlinfo.Caption = "Your attack missed!"
    bttlinfo.Visible = True
End If
End Sub
Sub cswing()
Dim sw As Integer
Dim npac As Integer
Randomize
sw = Int(Rnd * 20) + 1
npac = nac + 10
If (cac(v) + sw) >= npac Then
    catck
Else
    bttlinfo.Caption = "Your opponent missed!"
    bttlinfo.Visible = True
End If
End Sub
