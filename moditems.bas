Attribute VB_Name = "moditems"
'this has turned out to be more of a mod where everything
'goes into...well, i can understand it anyway.

Dim mac As Integer
Dim sac As Integer
Dim shp As Integer
Dim r1hp As Integer, r2hp As Integer
Dim r1mp As Integer, r2mp As Integer
Dim e1hp As Integer, e2hp As Integer
Dim e1mp As Integer, e2mp As Integer
Global nhp As Integer
Global nac As Integer
Global nbdmg As Integer
Dim lac As Integer
Dim hac As Integer
Dim hhp As Integer
Global nmhp As Integer
Global nmmp As Integer
Global monst As Integer
Global chp(0 To 2) As Integer
Global cdm(0 To 2) As Integer
Global cac(0 To 2) As Integer
Global cht(0 To 2) As Integer
Global ccr(0 To 2) As Integer
Global bst As Integer
Global wdmg As Integer


Sub items()

    
End Sub
Sub itmdesc()
darm1 = "Leather Armor.  AC: 1."
darm2 = "Studded leather armor.  AC: 2."
darm3 = "Chain mail armor.  AC: 3."
darm4 = "Scale mail armor.  AC: 4."
darm5 = "Banded plate mail armor.  AC: 5"
darm6 = "Full plate armor.  AC: 6"
dwpn1 = "Short sword.  Base damage: 2"
dwpn2 = "Long sword.  Base damage: 4"
dwpn3 = "Broad sword.  Base damage: 6"
dwpn4 = "Staff.  Base damage: 3.  +5 mp"
dshd1 = "Buckler.  AC: 1"
dshd2 = "Small shield.  AC: 1.  +5 hp"
dshd3 = "Large shield.  AC: 2.  +5 hp"
dshd4 = "Tower shield.  AC: 3.  +5 hp"
drng1 = "Ring.  +5 hp.  +5 mp"
dear1 = "Earring.  +5 hp.  +5 mp"
dleg1 = "leather leggings.  AC: 1"
dleg2 = "studded leather leggings.  AC: 2"
dleg3 = "Chain mail leggings.  AC: 3"
dleg4 = "Scale mail leggings. AC: 4"
dleg5 = "Banded plate mail leggings.  AC: 5"
dleg6 = "Full plate leggings.  AC: 6"
dhlm1 = "Leather cap.  AC: 1.  +5 hp"
dhlm2 = "Chain cap.  AC: 2.  +5 hp"
dhlm3 = "Plate helm.  AC: 3.  +5 hp"
End Sub
Sub stats()
Select Case m
    Case 1
        mac = 1
    Case 2
        mac = 2
    Case 3
        mac = 3
    Case 4
        mac = 4
    Case 5
        mac = 5
    Case 6
        mac = 6
End Select
If e1 = 16 Then
    e1hp = 5
    e1mp = 5
End If
If e2 = 16 Then
e2hp = 5
e2mp = 5
End If
If r1 = 15 Then
r1hp = 5
r1mp = 5
End If
If r2 = 15 Then
r2hp = 5
r2mp = 5
End If
Select Case h
    Case 23
        hac = 1 And hhp = 5
    Case 24
        hac = 2 And hhp = 5
    Case 25
        hac = 3 And hhp = 5
End Select
Select Case s
    Case 11
        sac = 1
    Case 12
        sac = 1 And shp = 5
    Case 13
        sac = 2 And shp = 5
    Case 14
        sac = 3 And shp = 5
End Select
Select Case w
    Case 7
        wdmg = 2
    Case 8
        wdmg = 4
    Case 9
        wdmg = 6
End Select
Select Case l
    Case 17
        lac = 1
    Case 18
        lac = 2
    Case 19
        lac = 3
    Case 20
        lac = 4
    Case 21
        lac = 5
    Case 22
        lac = 6
End Select
Select Case EXPer
    Case Is >= 20
            If level = 1 Then
                level = 2
                lpt = lpt + 3
                MsgBox "You have leveled up!", vbOKOnly, "Level up!"
            End If
    Case Is >= 50
            If level = 2 Then
                level = 3
                lpt = lpt + 3
                MsgBox "You have leveled up!", vbOKOnly, "Level up!"
            End If
    Case Is >= 90
            If level = 3 Then
                level = 4
                lpt = lpt + 3
                MsgBox "You have leveled up!", vbOKOnly, "Level up!"
            End If
End Select
mhp = Int(con / 2)
mmp = Int(inte / 2)
AC = Int(dex / 4)
bst = Int(str / 4)
nac = AC + lac + hac + mac + sac
bsd = Int(inte / 4)
nmhp = mhp + shp + hhp + e1hp + e2hp + r1hp + r2hp
nmmp = mmp + r1mp + r2mp + e1mp + e2mp
End Sub

Sub enctr()
Dim rnden As Integer
Dim enc As Boolean
Randomize
rnden = Int(Rnd * 100) + 1
If rnden <= 7 Then
        monst = 0
        enc = True
End If
If enc = True Then
enc = False
Load frmbttl
End If

End Sub

Sub Mons()
chp(0) = 12
cdm(0) = 3
cac(0) = 3
cht(0) = 3
ccr(0) = 5
chp(1) = 18
cdm(1) = 4
cac(1) = 4
cht(1) = 2
ccr(1) = 7
End Sub
