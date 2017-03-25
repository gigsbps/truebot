Attribute VB_Name = "Module3"
Public dnpcs As Scripting.Dictionary
Public ditems As Scripting.Dictionary
Public dskills As Scripting.Dictionary
Public dicExp1 As Scripting.Dictionary
Public dicExp2 As Scripting.Dictionary
Public dicExp3 As Scripting.Dictionary

Public LastSelectItem
Public LastSelectBPItem

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.number = 0)
   On Error GoTo 0
End Function





Sub SetNpc(npcid, npcname)
    Dim npc As Character

    Set npc = New Character
        npc.charname = npcname
        npc.uID = npcid
        dnpcs.Add npcid, npc
'    Set npc = Nothing

End Sub
Sub AddNpcSkill(npcid, skillid)
    Dim npc As Character
    Set npc = dnpcs.Item(npcid)
        npc.Skills.Add npc.Skills.Count, skillid
End Sub

Sub SetItems(itemid, itemname, itemtype, itemvalue, itemtype2, itemvalue2, itemdesc, eqposition)
    Dim itm As clsItems
    Set itm = New clsItems
    itm.itemid = itemid
    
    itm.itemname = itemname
    itm.itemtype = itemtype
    itm.itemvalue = itemvalue
    
    itm.itemtype2 = itemtype2
    itm.itemvalue2 = itemvalue2
    itm.itemdesc = itemdesc
    itm.eqposition = eqposition
    
    If (itemid >= 50001) And (itemid <= 50006) Then
        itm.itemtype = ""
        itm.itemtype2 = ""
    End If
    If Not ditems.Exists(itemid) Then
        ditems.Add itemid, itm
    End If
End Sub

Sub LoadNPCData()
Dim bArray() As Byte
Dim offset As Long
Dim datasize As Long
Dim filesize As Long
Dim filename As String
Dim foundDat As Boolean
    
    offset = 93 - 8
    datasize = 92 - 4
    foundDat = False
    If Form1.GameFolder <> "" Then
        If Dir(Form1.GameFolder & "\data\Npc.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) <> "" Then
            filename = Form1.GameFolder & "\data\Npc.Dat"
            foundDat = True
        End If
    End If
    If Not foundDat Then
        If Dir(App.Path & "\Npc.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
            If Dir("C:\Program Files\Asiasoft\TSOnline\data\Npc.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
                MsgBox "ไม่สามารถหาไฟล์ Npc.dat ได้", vbExclamation + vbOKOnly, "Error !!"
                Unload Form1
                End
            Else
                filename = "C:\Program Files\Asiasoft\TSOnline\data\Npc.Dat"
            End If
        Else
            filename = App.Path & "\Npc.Dat"
        End If
    End If
    bArray = ReadFile(filename, filesize)
    
    Index = 0
    Do While offset < filesize
        NPCidx = ((cInt2(bArray, offset + 15) Xor &H520A) Xor &H3) - 1
        NPCSkill1x = ((cInt2(bArray, offset + 65) Xor &H520A) Xor &H3) - 1
        NPCSkill2x = ((cInt2(bArray, offset + 67) Xor &H520A) Xor &H3) - 1
        NPCSkill3x = ((cInt2(bArray, offset + 69) Xor &H520A) Xor &H3) - 1
        
        NPCNamex = ""
        n = offset + 13
        Do While n > offset
            If bArray(n) = 0 Then Exit Do
            c = Chr(bArray(n))
            If Int(bArray(n)) >= 65 Then
            NPCNamex = NPCNamex & c
            End If
            n = n - 1
        Loop
        
        'Debug.Print NPCNamex
         
        
        Call SetNpc(NPCidx, NPCNamex)
        If NPCSkill1x > 0 Then
            Call AddNpcSkill(NPCidx, NPCSkill1x)
        End If
        If NPCSkill2x > 0 Then
            Call AddNpcSkill(NPCidx, NPCSkill2x)
        End If
        If NPCSkill3x > 0 Then
            Call AddNpcSkill(NPCidx, NPCSkill3x)
        End If
        offset = offset + datasize
    Loop
End Sub

