Attribute VB_Name = "Module2x"

Public Function GetNPCName(npcid) As String
    Dim npc As Character
    If dnpcs.Exists(npcid) Then
        Set npc = dnpcs.Item(npcid)
        GetNPCName = npc.charname
    Else
        GetNPCName = "Unknow NPC"
    End If
End Function

Public Function getNpcSkill(npcid)
    Dim npc As Character
    Set npc = dnpcs.Item(npcid)
    Set getNpcSkill = npc.Skills
End Function

Public Function getItemName(itemid)
    Dim itm As clsItems
    If Not ditems.Exists(itemid) Then
        getItemName = ""
        Exit Function
    End If
    Set itm = ditems.Item(itemid)
    If Not itm Is Nothing Then
        getItemName = itm.itemname
    End If
End Function
Public Function getItemType(itemid)
    Dim itm As clsItems
    If Not ditems.Exists(itemid) Then
        getItemType = ""
        Exit Function
    End If
    Set itm = ditems.Item(itemid)
    If Not itm Is Nothing Then
        getItemType = itm.itemtype
    End If
End Function
Public Function getItem(itemid)
    Dim itm As clsItems
    If Not ditems.Exists(itemid) Then
        Set getItem = Nothing
        Exit Function
    End If
    Set itm = ditems.Item(itemid)
    If Not itm Is Nothing Then
        Set getItem = itm
    End If
End Function


Public Function IsHP(itemid) As Boolean
    Dim itm As clsItems
       If Not ditems.Exists(itemid) Then
        IsHP = False
        Exit Function
    End If
    Set itm = ditems.Item(itemid)
    If itm.itemtype = "HP" And itm.itemvalue > 0 Then
        IsHP = True
        If (itemid >= 50001) And (itemid <= 50006) Then IsHP = False
    Else
        IsHP = False
    End If
    Exit Function
End Function
Public Function IsSP(itemid) As Boolean
    Dim itm As clsItems
       If Not ditems.Exists(itemid) Then
        IsSP = False
        Exit Function
    End If
    Set itm = ditems.Item(itemid)
    If itm.itemtype = "SP" And itm.itemvalue > 0 Then
        IsSP = True
        If (itemid >= 50001) And (itemid <= 50006) Then IsSP = False
    Else
        IsSP = False
    End If
    Exit Function
End Function

Public Function getSkillName(skillid)
'    Dim sk As clsSkill
    'Set sk = dskills.Item(skillid)
    'MsgBox dskills.Item(skillid).skillname
    If (skillid >= 10000) And (skillid <= 30000) Then
        getSkillName = dskills.Item(skillid).skillname
    Else
        getSkillName = "#" & skillid
    End If
End Function
Public Function getSkillId(skillname)
On Error Resume Next
 '   Dim sk As clsSkill
    For Each sk In dskills.Items
        If sk.skillname = skillname Then
            getSkillId = sk.skillid
            Exit Function
        End If
    Next
     getSkillId = 10000
End Function
Public Function getSkillSp(skillid)
     getSkillSp = dskills.Item(skillid).skillsp
End Function


Public Sub SetSkill(skillid, skillname, skillsp)
    Dim sk As clsSkill

    Set sk = New clsSkill
        sk.skillid = skillid
        sk.skillname = skillname
        sk.skillsp = skillsp
        dskills.Add skillid, sk
End Sub

Sub LoadSkillData()
Dim bArray() As Byte
Dim offset As Long
Dim datasize As Long
Dim filesize As Long
Dim filename As String
Dim foundDat As Boolean
    
    offset = 87
    datasize = 86
    foundDat = False
    If Form1.GameFolder <> "" Then
        If Dir(Form1.GameFolder & "\data\Skill.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) <> "" Then
            filename = Form1.GameFolder & "\data\Skill.Dat"
            foundDat = True
        End If
    End If
    If Not foundDat Then
        If Dir(App.Path & "\Skill.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
            If Dir("C:\Program Files\Asiasoft\TSOnline\data\Skill.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
                MsgBox "ไม่สามารถหาไฟล์ Skill.dat ได้", vbExclamation + vbOKOnly, "Error !!"
                Unload Form1
                End
            Else
                filename = "C:\Program Files\Asiasoft\TSOnline\data\Skill.Dat"
            End If
        Else
            filename = App.Path & "\Skill.Dat"
        End If
    End If
    bArray = ReadFile(filename, filesize)
    
    Index = 0
    Do While offset < filesize
        Skillidx = ((cInt2(bArray, offset + 21) Xor &H6EA4) Xor &H4) - 4
        SkillSpx = ((cInt2(bArray, offset + 23) Xor &H6EA4) Xor &H4) - 4
        
        SkillNamex = ""
        n = offset + &H13
        Do While n > offset
            If bArray(n) = 0 Then Exit Do
            c = Chr(bArray(n))
'
            SkillNamex = SkillNamex & c
            n = n - 1
        Loop
        Call SetSkill(Skillidx, SkillNamex, SkillSpx)
        offset = offset + datasize
    Loop
End Sub

