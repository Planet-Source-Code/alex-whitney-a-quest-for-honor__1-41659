Attribute VB_Name = "modSaveLoad"
    Global lost As Boolean
    Global fnf As Boolean
  Public SAVE_MapLoaded As String
  Public SAVE_Midi As String
  Public SAVE_OutsideMidi As String
  Public SAVE_SpeechLoaded As String
  Public SAVE_CharX As Integer
  Public SAVE_CharY As Integer
  Public SAVE_CharFacing As Integer
  Public SAVE_Wood As Integer
  Public SAVE_Coin As Integer
  Public SAVE_Magic As Integer
  Public SAVE_Ticket As Integer
  Public SAVE_Toast As Integer
  Public SAVE_SpellCut As String
  Public SAVE_SpellDestroy As String
  Public SAVE_level
  Public SAVE_str As Integer
  Public save_dex As Integer
  Public save_exp As Integer
  Public save_con As Integer
  Public save_inte As Integer
  Public save_h As Integer, save_e1 As Integer, save_e2 As Integer, save_m As Integer, save_w As Integer, save_s As Integer, save_r1 As Integer, save_r2 As Integer, save_l As Integer
  Global ar As Integer, wp As Integer, sh As Integer, le As Integer, rn As Integer, ea As Integer
  
    Sub dsv()
    On Error Resume Next
    kill (App.Path & "\SaveData.txt")
    End Sub
    
  

  

Public Sub SaveGame()
  SAVE_MapLoaded = MapLoaded
  SAVE_Midi = Midi
  SAVE_OutsideMidi = OutsideMidi
  SAVE_SpeechLoaded = SpeechLoaded
  SAVE_CharX = CharX
  SAVE_CharY = CharY
  SAVE_CharFacing = CharFacing
  SAVE_Wood = Wood
  SAVE_Coin = Coin
  SAVE_Magic = Magic
  SAVE_Ticket = Ticket
  SAVE_Toast = Toast
  SAVE_SpellCut = SpellCut
  SAVE_SpellDestroy = SpellDestroy
  SAVE_level = level
  SAVE_str = str: save_dex = dex: save_con = con: save_inte = inte
  save_h = h: save_e1 = e1: save_e2 = e2: save_r1 = r1: saver2 = r2: save_m = m: save_w = w: save_s = s: save_l = l
  save_exp = EXPer
  save_mmp = mmp
  save_mhp = mhp
  save_hp = hp
  save_mp = mp
  
  
  Open App.Path & "\SaveData.txt" For Output As 1
  Write #1, SAVE_MapLoaded, SAVE_Midi, SAVE_OutsideMidi, SAVE_SpeechLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_SpellCut, SAVE_SpellDestroy, SAVE_level, SAVE_str, save_dex, save_con, save_exp, SAVE_level, save_inte, save_h, save_e1, save_e2, save_m, save_w, save_s, save_r1, save_r2, save_l, save_mhp, save_mmp, save_hp, save_mp, lpt, osp(0), osp(1), osp(2), osp(3), dsp(0), dsp(1), dsp(2), dsp(3)
  Close
  MsgBox "Game Saved"
  
End Sub

Public Sub LoadGame()
  On Error GoTo err
  Open App.Path & "\SaveData.txt" For Input As 1
  Input #1, SAVE_MapLoaded, SAVE_Midi, SAVE_OutsideMidi, SAVE_SpeechLoaded, SAVE_CharX, SAVE_CharY, SAVE_CharFacing, SAVE_Wood, SAVE_Coin, SAVE_Magic, SAVE_Ticket, SAVE_Toast, SAVE_SpellCut, SAVE_SpellDestroy, level, str, dex, con, EXPer, level, inte, h, e1, e2, m, w, s, r1, r2, l, mhp, mmp, hp, mp, lpt, osp(0), osp(1), osp(2), osp(3), dsp(0), dsp(1), dsp(2), dsp(3)

    MapLoaded = SAVE_MapLoaded
    OutsideMidi = SAVE_OutsideMidi
    SpeechLoaded = SAVE_SpeechLoaded
    CharX = SAVE_CharX
    CharY = SAVE_CharY
    CharFacing = SAVE_CharFacing
    Wood = SAVE_Wood
    Coin = SAVE_Coin
    Magic = SAVE_Magic
    Ticket = SAVE_Ticket
    Toast = SAVE_Toast
    SpellCut = SAVE_SpellCut
    SpellDestroy = SAVE_SpellDestroy
    frmGame.Show
  Close
  Exit Sub
err:
  MsgBox "File not found.  Creating new game.", vbOKOnly, "Error"
  fnf = True
End Sub

