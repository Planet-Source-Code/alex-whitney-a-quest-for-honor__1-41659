Attribute VB_Name = "modVariables"


Public X As Integer
Public Y As Integer
Public Map(100) As String
Public PositionMap As String
Public MapLoaded As String

Public Midi As String
Public OutsideMidi As String

Public Message(100)
Public Speaker As String
Public NewLine As String
Public SpeechLoaded As String

Public PassToNext As Integer

Public CharX As Integer
Public CharY As Integer
Public CharFacing As Integer

Public FloorMode As String

'Game Variables
Public SpellCut As Boolean
Public SpellDestroy As Boolean

Public Item As Integer
Public Wood As Integer
Public Coin As Integer
Public Magic As Integer
Public Ticket As Integer
Public Toast As Integer

Public ItemHouseFound1 As Boolean
Public ItemHouseFound2 As Boolean
Public ItemHouseFound3 As Boolean

Public ItemTownFound1 As Boolean
Public ItemTownFound2 As Boolean
Public ItemTownFound3 As Boolean

'******************
'Letter Definition.
'******************

'G = Grass
'B = Bush
'0 = Sign  'Cero
'Q = Top Out Left Border
'A = Bottom Out Left Border
'W = Top Out Right Border
'S = Bottom Out Right Border
'Z = Inn Left Border
'X = Inn Right Border
'z = Inn Top Border
'x = Inn Bottom Border
'- = Water
'_ = Nothing
'For DX!
Public DX As New DirectX8
Public DI As DirectInput8
Public Keyboard As DirectInputDevice8
Public Mouse As DirectInputDevice8
Public KeyboardState As DIKEYBOARDSTATE
Public MouseState As DIMOUSESTATE
Public keyboard2state As DIKEYBOARDSTATE
'stats
Public str As Integer
Public dex As Integer
Public con As Integer
Public wis As Integer
Public inte As Integer
Public EXPer As Integer
Public hp As Integer
Global mhp As Integer
Public mp As Integer
Global mmp As Integer
Public AC As Integer
Public arm(1 To 6) As Boolean
Public wpn(1 To 4) As Boolean
Public shd(1 To 4) As Boolean
Public rng(1 To 2) As Boolean
Public ear(1 To 2) As Boolean
Public leg(1 To 6) As Boolean
Public hlm(1 To 4) As Boolean
Public level As Integer
Global itm(0 To 25) As Object
Public darm1 'global vars for item desc...
Public darm2 'could be MUCH more efficient, but
Public darm3 'to redo this would mean i would have to change
Public darm4 'a LOT of code...faster coding on my part at
Public darm5 'this pt to just kepe going with it
Public darm6
Public dwpn1
Public dwpn2
Public dwpn3
Public dwpn4
Public dshd1
Public dshd2
Public dshd3
Public dshd4
Public drng1
Public dear1
Public dleg1
Public dleg2
Public dleg3
Public dleg4
Public dleg5
Public dleg6
Public dhlm1
Public dhlm2
Public dhlm3
Global holding As Object
Global h As Integer
Global e1 As Integer
Global e2 As Integer
Global m As Integer
Global w As Integer
Global s As Integer
Global r1 As Integer
Global r2 As Integer
Global l As Integer
Global osp(0 To 3) As Boolean
Global dsp(0 To 3) As Boolean
Global lpt As Integer
