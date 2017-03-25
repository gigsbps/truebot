VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tspacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Character"
Attribute VB_Ext_KEY = "Member1" ,"NPCSCombat"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents mvarobj As Winsock
Attribute mvarobj.VB_VarHelpID = -1

Public lastNPCtalk As Integer

Public Event PatchIncorrect()
Public Event Connected()
Public Event ReadyForLogin()
Public Event Closed()
Public Event ConnectionError(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
Public Event PlayerOnline(ByVal onlineDatetime As Date)
Public Event AppearCurrentOnlinePlayers(ByVal objPlayerCharacter As Character)
Public Event AppearOnlinePlayers(ByVal objPlayerCharacter As Character)
Public Event PlayerOffline(ByVal playerid As Long)
Public Event PlayerAppearInMap(ByVal playerid As Long, x, y)
Public Event PlayerLeaveMap(ByVal playerid As Long, mapid, x, y)
Public Event onSendClickOnNpc(ByVal nid As Long)
Public Event DoSelectPartner(ByVal partnerid As Long)
Public Event DoUnSelectPartner()
Public Event MyAttack()
Public Event PartnerAttack()
Public Event DuplicateLogin()
Public Event TsError()
Public Event RecvPartnerLists(ByVal p As Character)
Public Event InitPlayerStatus()
Public Event NpcDialog(ByVal DialogId As Long)
Public Event ChatMessage(typeid, Msg, sender)
Public Event onBattleStoped()
Public Event onBattleStarted()
Public Event onOpenCombat()
Public Event onWaitngAttack()
Public Event warpFinish()
Public Event onNPCAppear(ByVal npcmapid As Integer, ByVal x As Long, ByVal y As Long)
Public Event AppearAnotherCombat(ByVal playerid As Long)
Public Event WaitngForAcceptParty(ByVal playerid As Long)
Public Event doAcceptParty(ByVal partyid As Long)
Public Event doNotEnoughSlot(ByVal itemid As Long, ByVal n As Integer)
Public Event doNotEnoughBackPackSlot(ByVal itemid As Long, ByVal n As Integer)
Public Event InvalidLicence()
Public Event ValidLicence()
Public Event NpcDialogMenu(ByVal DialogId As Long)
Public Event NpcWalkThenDialog(ByVal DialogId As Long)
Public Event CombatSceneDialog(ByVal DialogId As Long)
Public Event ResponseAnswer()
Public Event RecvNPCCombat()
Public Event FinishBattle(ByVal uID As Long)
Public Event RequestPartyCancel(ByVal uID As Long)
Public Event RequestPartyFalse(ByVal uID As Long)
Public Event RequestPartyAcceptFrom(ByVal uID As Long)
Public Event Connecting()
Public Event FinishAnswerFuckGod()
Public Event RecvMoney(ByVal money As Long)
Public Event RecvItemFrom(ByVal uID As Long, ByVal itemid As Long, ByVal n As Integer)
Public Event SendItemSuccess(ByVal uID As Long, ByVal itemid As Long, ByVal n As Integer)
Public Event RecvQuestion()
Public Event InitInventoryList()
Public Event InitBackPackList()
Public Event onDamage()
Public Event onChangeStatus()
Public Event RecvDropItems(ByVal itemid As Long, ByVal num As Integer)
Public Event RecvBackPackItems(ByVal itemid As Long, ByVal num As Integer)
Public Event InventoryChange()
Public Event BackPackChange()
Public Event onWarp(ByVal uID As Long, ByVal mapid As Long, ByVal warpid As Integer)
Public Event onState1()
Public Event onSetsena(ByVal uID As Long)
Public Event onRequestSleep(ByVal price As Long)
Public Event LoginFail()
Public Event Loginok()
Public Event onSendAttack(ByVal fr As Integer, ByVal fc As Integer, ByVal tr As Integer, ByVal tc As Integer, ByVal sk As Long)
Public Event onEquipment(ByVal slotno As Integer)
Public Event on140B()

Public Event ArmyListOnline(ByVal uID As Long)
Public Event ArmyListOffline(ByVal uID As Long)
Public Event InitArmyList()
Public Event PartyUpdateHP(ByVal uID As Long, ByVal HP)
Public Event PartyUpdateSP(ByVal uID As Long, ByVal SP)


Private mvarCharacter As Character
Public ol As Dictionary
Public MyPartners As Dictionary
Public MyParties As Dictionary
Public oNPCCombat As Dictionary
Public oNPCCombayDmg As Dictionary
Public MyItems As Dictionary 'local copy
Public BackPack As Dictionary 'local copy
Public CurrentPartner As Character
Dim dmgsource As DamageInfo
Dim DmgTarget As DamageTarget
Private mvarNPCSCombat As NPCSCombat
Private npcscom As New NPCSCombat
Private partner As New Character
Public ConnectionState
Private mvarDialogId As Long 'local copy

Private tep
Dim nd As NPCSCombat
Private dBuffer As String
Private mvarLastMsgType As Integer 'local copy
Private mvarLastMsgSenderId As Long 'local copy
Private mvarLastMessage As String 'local copy
Private mvarLastPlayerLogoff As Long 'local copy
Private mvarLastPlayerOnline As Long 'local copy
Private mvarLastQuestion As String 'local copy
Public LastAnswers As Dictionary 'local copy
Private mvarInc As Inv
Private mvarBackc As Inv
Private mvarLastRequestPartyId As Long 'local copy


Private mvarLastDamageRowPosition As Integer 'local copy
Private mvarLastDamageColPosition As Integer 'local copy
Private mvarLastDamageSkillId As Long 'local copy
Private mvarLastWarpId As Integer 'local copy
Private mvarLastLeaderWarpId As Integer 'local copy
Private mvarLastMapId As Long 'local copy


Public WarpState As Boolean
Public fightingstate As Boolean
Public godquestion As Boolean

Dim requesttowarp As Boolean
Dim testid As Long


Public partyLeader As Boolean

Public pversion As Integer


Public IsHorse As Boolean





Dim xxxx
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent PlayerWalk[(arg1, arg2, ... , argn)]
Public Event PlayerWalk(ByVal playerid As Long, ByVal direction As Integer, ByVal x As Long, ByVal y As Long)
'local variable(s) to hold property value(s)
Private mvarCurrentParty As Long 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent PartyStop[(arg1, arg2, ... , argn)]
Public Event PartyStop(ByVal playerid As Long)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onWalk[(arg1, arg2, ... , argn)]
Public Event onWalk(x As Variant, y As Variant)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onSales[(arg1, arg2, ... , argn)]
Public Event onSales(itemid, num, money)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onGotGhost[(arg1, arg2, ... , argn)]

Public Event onGotGhost()
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onGotLuckyGod[(arg1, arg2, ... , argn)]
Public Event onGotLuckyGod()
'local variable(s) to hold property value(s)
Private mvarLastResponseAnswer As String 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onPlayerWalk[(arg1, arg2, ... , argn)]
Public Event onPlayerWalk(x As Long, y As Long, playerid As Long)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onAnswerRight[(arg1, arg2, ... , argn)]
Public Event onAnswerRight(ByVal Question As String, ByVal answer As String)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onAnswerWrong[(arg1, arg2, ... , argn)]
Public Event onAnswerWrong(ByVal Question As String, ByVal answer As String)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent odEattem[(arg1, arg2, ... , argn)]
Public Event odEattem(ByVal slot As Integer, ByVal n As Integer, ByVal pos As Integer)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onUpdateValue[(arg1, arg2, ... , argn)]
Public Event onUpdateValue()
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onUpdateTexp[(arg1, arg2, ... , argn)]
Public Event onUpdateTexp()






Public Property Let LastResponseAnswer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastResponseAnswer = 5
    mvarLastResponseAnswer = vData
End Property


Public Property Get LastResponseAnswer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastResponseAnswer
    LastResponseAnswer = mvarLastResponseAnswer
End Property








Public Property Let CurrentParty(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentParty = 5
    mvarCurrentParty = vData
End Property


Public Property Get CurrentParty() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentParty
    CurrentParty = mvarCurrentParty
End Property








Public Sub delay(millisec)
    Sleep (millisec)
End Sub









Public Property Let lastMapID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastWarpId = 5
    mvarLastMapId = vData
End Property


Public Property Get lastMapID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastWarpId
    lastMapID = mvarLastMapId
End Property



Public Property Let LastWarpId(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastWarpId = 5
    mvarLastWarpId = vData
End Property


Public Property Get LastWarpId() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastWarpId
    LastWarpId = mvarLastWarpId
End Property

Public Property Let LastLeaderWarpId(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastWarpId = 5
    mvarLastLeaderWarpId = vData
End Property


Public Property Get LastLeaderWarpId() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastWarpId
    LastLeaderWarpId = mvarLastLeaderWarpId
End Property



Public Property Let LastDamageSkillId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastDamageSkillId = 5
    mvarLastDamageSkillId = vData
End Property


Public Property Get LastDamageSkillId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastDamageSkillId
    LastDamageSkillId = mvarLastDamageSkillId
End Property



Public Property Let LastDamageColPosition(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastDamageColPosition = 5
    mvarLastDamageColPosition = vData
End Property


Public Property Get LastDamageColPosition() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastDamageColPosition
    LastDamageColPosition = mvarLastDamageColPosition
End Property



Public Property Let LastDamageRowPosition(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastDamageRowPosition = 5
    mvarLastDamageRowPosition = vData
End Property


Public Property Get LastDamageRowPosition() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastDamageRowPosition
    LastDamageRowPosition = mvarLastDamageRowPosition
End Property







Public Property Let LastRequestPartyId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastRequestPartyId = 5
    mvarLastRequestPartyId = vData
End Property


Public Property Get LastRequestPartyId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastRequestPartyId
    LastRequestPartyId = mvarLastRequestPartyId
End Property








Public Property Let LastQuestion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastQuestion = 5
    mvarLastQuestion = vData
End Property


Public Property Get LastQuestion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastQuestion
    LastQuestion = mvarLastQuestion
End Property



Public Property Let LastPlayerOnline(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastPlayerOnline = 5
    mvarLastPlayerOnline = vData
End Property


Public Property Get LastPlayerOnline() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastPlayerOnline
    LastPlayerOnline = mvarLastPlayerOnline
End Property



Public Property Let LastPlayerLogoff(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastPlayerLogoff = 5
    mvarLastPlayerLogoff = vData
End Property


Public Property Get LastPlayerLogoff() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastPlayerLogoff
    LastPlayerLogoff = mvarLastPlayerLogoff
End Property



Public Property Let LastMessage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastMessage = 5
    mvarLastMessage = vData
End Property


Public Property Get LastMessage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastMessage
    LastMessage = mvarLastMessage
End Property



Public Property Let LastMsgSenderId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastMsgSenderId = 5
    mvarLastMsgSenderId = vData
End Property


Public Property Get LastMsgSenderId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastMsgSenderId
    LastMsgSenderId = mvarLastMsgSenderId
End Property



Public Property Let LastMsgType(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastMsgType = 5
    mvarLastMsgType = vData
End Property


Public Property Get LastMsgType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastMsgType
    LastMsgType = mvarLastMsgType
End Property

Private Sub senddata(data() As Byte)
On Error Resume Next
    If mvarobj.state <> 0 Then
        mvarobj.senddata data
        DoEvents
    End If
End Sub
Private Sub senddataX(data() As Byte)
On Error Resume Next
    If mvarobj.state <> 0 Then
        mvarobj.senddata data
    End If
End Sub

Public Sub JamCombat(uID)
'F4 44 09 00 0B 02 05 A8 3A 04 00 00 00
   senddata MakePacket("F4 44 09 00 0B 02 05 " & n2h(uID, 4) & " 00 00")
        DoEvents
End Sub

Public Sub doSleep()
   senddata MakePacket("F4 44 02 00 1F 01")
   DoEvents
End Sub

Public Sub CheckAirTime()
    senddata MakePacket("F4 44 02 00 22 01")
    DoEvents
End Sub

'
Public Sub SendAction(acid)
        senddata MakePacket("F4 44 03 00 20 02" & n2h(acid, 1))
        DoEvents
End Sub
'
'Public Sub Logout()
'
'End Sub
Public Sub DealOpen(playerid)
     'F4 44 06 00 19 01 32 3A 07 00
        senddata MakePacket("F4 44 06 00 19 01 " & n2h(playerid, 4))
        DoEvents
End Sub

Public Sub Horse()
On Error Resume Next
    For Each Order In Me.MyPartners.Keys
        If (Me.MyPartners(Order).uID > "18000") And (Me.MyPartners(Order).uID < "19000") Then
            senddata MakePacket("F4 44 06 00 0F 04 " & n2h(Me.MyPartners(Order).uID, 2) & "  00 00")
            IsHorse = True
            DoEvents
            Exit For
        End If
    Next
End Sub

Public Sub SendItemTo(uID, slot, num)
On Error Resume Next
    For Each Order In Me.MyPartners.Keys
        If Me.MyPartners(Order).uID = "22002" _
        Or Me.MyPartners(Order).uID = "22003" _
        Or Me.MyPartners(Order).uID = "22004" _
        Or Me.MyPartners(Order).uID = "22005" _
        Or Me.MyPartners(Order).uID = "22028" _
        Or Me.MyPartners(Order).uID = "22029" _
        Or Me.MyPartners(Order).uID = "22039" _
        Or Me.MyPartners(Order).uID = "22048" _
        Or Me.MyPartners(Order).uID = "41187" _
        Then
            senddata MakePacket("F4 44 14 00 19 14 " & n2h(Me.MyPartners(Order).uID, 2) & "00 00" & n2h(uID, 4) & n2h(slot, 1) & n2h(num, 1) & "00 00 00 00 00 00 00 00")
            DoEvents
            Exit For
        End If
    Next
End Sub

Public Function Warp(warpid) As Boolean
                                 'F4 44 04 00 14 04 02 00
                                 'F4 44 04 00 14 08 02 00
     warpIdhex = n2h(warpid, 2)
     
        requesttowarp = True
        Me.LastWarpId = warpid
        senddata MakePacket("F4 44 04 00 14 04 " & warpIdhex)
        WarpState = True
        DoEvents
End Function
Public Sub AcceptParty(uID)
        partyLeader = True
        senddata MakePacket("F4 44 07 00 0D 03 01 " & n2h(uID, 4))
        DoEvents
End Sub

'Public Sub SendItem()
'End Sub
Public Sub UnHorse()
'F4 44 02 00 0F 05
        IsHorse = False
        senddata MakePacket("F4 44 02 00 0F 05")
        DoEvents
End Sub
Public Sub UnSelectPartner()
        senddata MakePacket("F4 44 02 00 13 02")
        DoEvents
End Sub

Public Sub SelectPartner(npcid)
        If npcid = "18001" _
        Or npcid = "18002" _
        Or npcid = "18003" _
        Or npcid = "18004" _
        Or npcid = "18005" _
        Or npcid = "18006" _
        Or npcid = "18007" _
        Or npcid = "18008" _
        Or npcid = "18009" _
        Or npcid = "18010" _
        Or npcid = "18011" _
        Or npcid = "18012" _
        Or npcid = "18013" _
        Or npcid = "18014" _
        Or npcid = "18015" _
        Then
            If IsHorse = True Then
                Exit Sub
            End If
        End If
        senddata MakePacket("F4 44 06 00 13 01 " & n2h(npcid, 2) & " 00 00")
        DoEvents
End Sub
Public Sub Equipment(slot)
    'F4 44 03 00 17 0B 19
        senddata MakePacket("F4 44 03 00 17 0B " & n2h(slot, 1))
        DoEvents
        'Form1.updateinv
End Sub

Public Sub UnEquip(slot, indx)
    'F4 44 04 00 17 0C 04 01
        senddata MakePacket("F4 44 04 00 17 0C " & n2h(slot, 1) & n2h(indx, 1))
        DoEvents
        'Form1.updateinv
End Sub

Public Sub LeaverParty(uID)
    senddata MakePacket("F4 44 06 00 0D 04 " & n2h(uID, 4))
    DoEvents
End Sub

Public Sub EatItem(slot, n, pos)
    'F4 44 06 00 17 0F 13 01 00 00
    'HexStr(hexnum)
    'F4 44 06 00 17 0F 12 01 01 00
    'F4 44 06 00 17 0F 12 01 02 00
    senddata MakePacket("F4 44 06 00 17 0F " & n2h(slot, 1) & n2h(n, 1) & " " & n2h(pos, 1) & " 00")
    DoEvents
End Sub
Public Sub EatItemForAuto(slot, n, pos)
    If fightingstate = True Then Exit Sub
    senddata MakePacket("F4 44 06 00 17 0F " & n2h(slot, 1) & n2h(n, 1) & " " & n2h(pos, 1) & " 00")
    DoEvents
End Sub
Public Sub Heal(skillid, uID)
'F4 44 0B 00 17 14 02 FC 2A 02 00 DC E8 05 00
    senddata MakePacket("F4 44 0B 00 17 14 02 " & n2h(skillid, 2) & " 02 00 " & n2h(uID, 4))
    DoEvents
End Sub
Public Sub HealPartner(skillid, Order)
'F4 44 0B 00 17 14 02 FC 2A 04 01 92 A0 00 00
'F4 44 0B 00 17 14 02 FC 2A 04 02 0D 56 00 00
'F4 44 0B 00 17 14 02 FC 2A 04 01 75 A0 00 00
'F4 44 0B 00 17 14 02 FC 2A 04 02 F3 55 00 00
    senddata MakePacket("F4 44 0B 00 17 14 02 " & n2h(skillid, 2) & " 04 " & n2h(Order, 1) & n2h(Me.MyPartners(Order - 1).uID, 4))
    DoEvents
End Sub
Public Sub UseItemForPlayer(itemid, uID)
'F4 44 0B 00 17 14 02 FC 2A 02 00 DC E8 05 00
'F4 44 0B 00 17 14 01 BA 69 04 01 2D 2B 00 00
    senddata MakePacket("F4 44 0B 00 17 14 01 " & n2h(itemid, 2) & " 02 00 " & n2h(uID, 4))
    DoEvents
End Sub
Public Sub UseItemForPartner(itemid, npcid)
'F4 44 0B 00 17 14 02 FC 2A 02 00 DC E8 05 00
'F4 44 0B 00 17 14 01 BA 69 04 01 2D 2B 00 00
    senddata MakePacket("F4 44 0B 00 17 14 01 " & n2h(itemid, 2) & " 04 01 " & n2h(npcid, 4))
    DoEvents
End Sub
Public Sub UseItem(slot, n, pos)
    senddata MakePacket("F4 44 06 00 17 0F " & n2h(slot, 1) & n2h(n, 1) & " " & n2h(pos, 1) & " 00")
    DoEvents
End Sub


Public Sub DropItem(slot, num)
    'F4 44 04 00 17 03 0E 01
    'F4 44 04 00 17 03 09 01
    senddata MakePacket("F4 44 04 00 17 03 " & n2h(slot, 1) & n2h(num, 1))
    DoEvents
End Sub

Public Sub Contribute(money, slot)
'F4 44 07 00 27 0F 00 00 00 00 04
    senddata MakePacket("F4 44 07 00 27 0F " & n2h(money, 4) & n2h(slot, 1))
    DoEvents
 
End Sub


Public Sub Chat(ByVal ctype As Integer, ByVal Message As String, Optional uID As Long)
  ' F4 44 27 00 02 0B 00 00 00 00 C2 D4 B9 B4 D5 B5 E9 CD B9 C3 D1 BA CA D9 E8 E2 C5 A1 E1 CB E8 A7 20 54 53 20 4F 6E 6C 69 6E 65 31
    packet = "F444"
    If uID = 0 Then
        uID = Me.Character.uID
    End If
   
   If ctype = 3 Then
        If uID = 0 Then
         Exit Sub
        Else
            data = "02" & n2h(ctype, 1) & n2h(uID, 4) & a2hstr(Message)
            psize = Int(Len(data) / 2)
            
            packet = packet & n2h(psize, 2) & data
            
        End If
   Else
            data = "02" & n2h(ctype, 1) & a2hstr(Message)
            psize = Int(Len(data) / 2)
            
            packet = packet & n2h(psize, 2) & data
    
   End If
   
   
   
    'MsgBox packet
    senddata MakePacket(packet)
    DoEvents
    'F4 44 0D 00 02 03 A8 3A 04 00 C3 D1 A1 E2 C2 B9 D0
End Sub

Public Sub PKUser(ByVal uID As Long)
    'For license ID
    ' F4 44 09 00 0B 02 02 A8 12 01 00 00 00
    'packet = "F44409000B0202" & n2h(uID, 4) & "0000"
    'senddata MakePacket(packet)
    'DoEvents
End Sub

Public Sub Shop(ShopName As String _
 , item1 As Scripting.Dictionary _
)
'F4 44 0C 00 17 1E 03 78 78 78 01 0B 14 00 00 00
'F4 44 12 00 17 1E 04 78 78 78 78 01 0B 05 00 00 00 0E 06 00 00 00
      datastr = "17 1E"
      datalen = 0
      NameLen = Len(Trim(ShopName))
      datastr = "17 1E" & n2h(NameLen, 1)
      datastr = datastr & a2hstr(ShopName) & "01"
      datalen = 3 + NameLen + 1
      For i = 1 To 25
        If item1.Item(i) <> 0 Then
          datalen = datalen + 5
          datastr = datastr & n2h(i, 1) & n2h(item1.Item(i), 4)
        End If
      Next
      senddata MakePacket("F4 44" & n2h(datalen, 2) & datastr)
      DoEvents
End Sub
Public Sub CloseShop()
 'F4 44 02 00 17 1F
      senddata MakePacket("F4 44 02 00 17 1F")
      DoEvents
End Sub





Private Sub Class_Initialize()
   ' VB.Licenses.Add "truebot.tspacket", "123456"
    Set mvarobj = Form1.Winsock1
    Set mvarCharacter = New Character
    Set CurrentPartner = New Character
    Set ol = New Dictionary
    Set MyPartners = New Dictionary
    Set oNPCCombat = New Dictionary
    Set LastAnswers = New Dictionary
    Set mvarNPCSCombat = New NPCSCombat
    Set MyParties = New Dictionary
    
    Set MyItems = New Dictionary
    Set BackPack = New Dictionary
    
    For i = 1 To 25
     Set mvarInc = New Inv
         mvarInc.itemid = 0
         mvarInc.slot = i
         mvarInc.num = 0
         Me.MyItems.Add i, mvarInc
    Next
    For i = 1 To 25
     Set mvarBackc = New Inv
         mvarBackc.itemid = 0
         mvarBackc.slot = i
         mvarBackc.num = 0
         Me.BackPack.Add i, mvarBackc
    Next
    
        Me.ConnectionState = 0
        dBuffer = ""
        tep = 0
        WarpState = False
        requesttowarp = False
        RequestToWork = False
        fightingstate = False
        godquestion = False
        
         testid = 661976
         
    Me.Character.gold = 0
         partyLeader = False
        
End Sub


Public Property Let DialogId(ByVal vData As Long)
    mvarDialogId = vData
End Property
Public Property Get DialogId() As Long
    DialogId = mvarDialogId
End Property
Public Property Get NPCSCombat() As NPCSCombat
    Set NPCSCombat = mvarNPCSCombat
End Property
Public Property Set NPCSCombat(vData As NPCSCombat)
    Set mvarNPCSCombat = vData
End Property
Public Property Get Character() As Character
    Set Character = mvarCharacter
End Property
Public Property Set Character(vData As Character)
    Set mvarCharacter = vData
End Property
Private Sub Class_Terminate()
  Set mvarNPCSCombat = Nothing
    Set mvarCharacter = Nothing
End Sub


Public Sub CloseConnection()
    If mvarobj.state <> 0 Then
        mvarobj.Close
    End If
End Sub

Public Sub ConnectServer(ipAddress As String, port As Long)
'    If Left(ipAddress, 7) <> "203.144" Then
'        Form1.AppendDisplay "This Truebot was develop for Thai-TS ONLY !!", vbRed
'        Me.Disconect
'    Else
'        If mvarobj.state = 0 Then
'            mvarobj.Connect ipAddress, port
'        End If
'        DoEvents
'        RaiseEvent Connecting
'    End If
    If mvarobj.state = 0 Then
            mvarobj.Connect ipAddress, port
    End If
    DoEvents
    RaiseEvent Connecting
End Sub
Public Sub SendEnd()
'F4 44 02 00 14 06
    senddata MakePacket("F4 44 02 00 14 06")
    DoEvents
End Sub
Public Sub RequestLogin()
    If mvarobj.state = sckClosed Then
        Exit Sub
    End If
    senddata MakePacket("F4 44 01 00 00")
    DoEvents
End Sub

'F444 0600 10 01 01 000000
'F444 0600 10 01 02000000
'F444 0600 10 01 03000000
'F444 0600 10 01 04000000
Public Sub answer(n)
Dim c
    For Each c In Me.LastAnswers.Keys
        If Me.LastAnswers(c) = n Then
           Me.LastResponseAnswer = c
        End If
    Next
    senddata MakePacket("F444 0600 10 01" + n2h(n, 1) + "000000")
    DoEvents
End Sub

Public Sub SelectChoice(n)
    senddata MakePacket("F444 0300 14 09 " + n2h(&H1D + n, 1))
    DoEvents
End Sub

Public Sub CancelQuest()
    senddata MakePacket("F444 0200 1409 ")
    DoEvents
    senddata MakePacket("F444 0400 14 01 63 00")
    DoEvents
    senddata MakePacket("F444 0400 14 01 63 00")
    DoEvents
End Sub

Public Sub sena(uID)
    senddata MakePacket("F444 0600 0D 05" + n2h(uID, 4))
    DoEvents
    
End Sub

Public Sub MixItem(slot1, num1, slot2, num2, slot3, num3)
'F4 44 08 00 17 0E 03 01 02 01 01 01
    senddata MakePacket("F4 44 08 00 17 0E " + n2h(slot1, 1) + " " + n2h(num1, 1) + n2h(slot2, 1) + " " + n2h(num2, 1) + n2h(slot3, 1) + " " + n2h(num3, 1))
    DoEvents
End Sub


Public Sub BackPackToInven(slotno)
'F4 44 03 00 17 25 0F  << เอาช่อง 15 มา inven
    senddata MakePacket("F4 44 03 00 17 25" + n2h(slotno, 1))
    DoEvents
End Sub

Public Sub InvenToBackPack(slotno)
'F4 44 03 00 17 24 0D << เอาช่อง 13 ใส่เป้
    senddata MakePacket("F4 44 03 00 17 24" + n2h(slotno, 1))
    DoEvents
End Sub

Public Sub Login(uID, passwd)

'F4 44 13 00 01 09 37 8B 02 00 74 73 7B 00 30 39 36 33 31 33 31 32 35
'F4 44 13 00 01 09 90 51 01 00 74 6E 7B 00 31 32 33 34 35 36 37 38 39
'F4 44 10 00 01 06 32 3A 07 00 74 73 7C 00 79 6F 6C 61 6E 64
logID = uID
logPass = passwd
    If mvarobj.state = sckClosed Then
        Exit Sub
    End If
        usertype = LCase(Mid(uID, 1, 2))
    
        tsid = Mid(uID, 3)
'
        Me.Character.uID = CLng(tsid)
        hexid = n2h(tsid, 4)
        
      '  Text3.Text = hexid & vbNewLine
        header = "F4 44"
        pckString = header
   '      Text3.Text = Text3.Text & pckstring & vbNewLine
        'version = "74 73 79 00"
        '74 73 7D 00 - patch 31/05/2548
        If usertype = "ts" Then
            uflag = "73"
        ElseIf usertype = "tn" Then
            uflag = "6E"
        End If
        'Version = "74 " & uflag & " 7E 00"
        'Version = "74 " & uflag & " 7F 00"
        'Version = "74 " & uflag & " 80 00"
        'Version = "74 " & uflag & " 82 00"
        'Version = "74 " & uflag & " 83 00"
        Version = "74 " & uflag & n2h(pversion, 1) & "  00"
        Version = "6D 62 8F 00"
        Version = "6D 62 B9 00"
        
        
        
        passlen = Len(Trim(passwd))
        
        DataLength = Right("00" & Hex(10 + passlen), 2) & "00"
        cmd = "01"
        hpass = ""
        For i = 1 To passlen
            ch = Hex(Asc(Mid(passwd, i, 1)))
            hpass = hpass & ch
        Next
        pckString = pckString & DataLength & cmd & n2h(passlen, 1) & hexid & Version & hpass
       ' SendLogin = MakePacket(pckstring)
    senddata MakePacket(pckString)
    DoEvents
End Sub

Public Sub Sale(slot, n)
'F4 44 04 00 1B 02 12 01
'F4 44 04 00 1B 02 07 01
    senddata MakePacket("F4 44 04 00 1B 02 " & n2h(slot, 1) & n2h(n, 1))
    DoEvents
End Sub
Public Sub MoveItem(slot1, slot2, num)
'F4 44 05 00 17 0A 05 01 04
    senddata MakePacket("F4 44 05 00 17 0A " & n2h(slot1, 1) & " " & n2h(num, 1) & " " & n2h(slot2, 1) & "")
    DoEvents
End Sub

Public Sub ClickOnNPC(nid As Long)
    Dim knid As String
    knid = Right("00" & Hex(nid), 2)
    pck = "F444 0400 14 01 " & knid & " 00"
    senddata MakePacket(pck)
    DoEvents
    RaiseEvent onSendClickOnNpc(nid)
End Sub

Public Sub Walk(x, y)
'F4 44 09 00
'06 01
'00
'46 02
'3B 01
'6B 36
    curX = Me.Character.x
    curY = Me.Character.y
'
'      1    0    7
'      2    X    6
'      3    4    5
'
    If x = Me.Character.x And y < Me.Character.y Then
        direction = 0
    ElseIf x < Me.Character.x And y < Me.Character.y Then
        direction = 1
    ElseIf x < Me.Character.x And y = Me.Character.y Then
        direction = 2
    ElseIf x < Me.Character.x And y > Me.Character.y Then
        direction = 3
    ElseIf x = Me.Character.x And y > Me.Character.y Then
        direction = 4
    ElseIf x > Me.Character.x And y > Me.Character.y Then
        direction = 5
    ElseIf x > Me.Character.x And y = Me.Character.y Then
        direction = 6
    ElseIf x > Me.Character.x And y < Me.Character.y Then
        direction = 7
    End If
    direction = Right("00" & Hex(direction), 2)
    xhex = Right("0000" & Hex(x), 4)
    hx = Left(xhex, 2)
    lx = Right(xhex, 2)
    xhex = lx & hx
    
    yhex = Right("0000" & Hex(y), 4)
    hx = Left(yhex, 2)
    lx = Right(yhex, 2)
    yhex = lx & hx
    Randomize
    z = Right("00" & Hex(Int((255 * Rnd) + 200)), 2)
    pck = "F4 44 09 00 06 01" & direction & xhex & yhex & n2h(Me.Character.level, 1) & z
    
     Me.Character.x = x
     Me.Character.y = y
   
    
    senddata MakePacket(pck)
    RaiseEvent onWalk(x, y)
    DoEvents
End Sub

Public Sub SendAttack(fr, fc, tr, tc, skill)

'F4 44 0A 00 32 02 03 02 02 02 BA 69 0A A0
'        If (tr > 3) Or (tr < 0) Then
'            tr = Int(4 * Rnd)
'        End If
'
'        If (tc > 3) Or (tc < 0) Then
'            tc = Int(4 * Rnd)
'        End If
        
        fr = Right("00" & fr, 2)
        fc = Right("00" & fc, 2)
        tr = Right("00" & tr, 2)
        tc = Right("00" & tc, 2)
        
        skillid = n2h(skill, 2)
        senddata MakePacket("F4 44 0A 00 32 01 " & fr & " " & fc & " " & tr & " " & tc & " " & skillid & " 0A A0")
        DoEvents
End Sub
Public Sub SendAttackItem(fr, fc, tr, tc, itemid)

'F4 44 0A 00 32 02 03 02 02 02 BA 69 0A A0
        fr = Right("00" & fr, 2)
        fc = Right("00" & fc, 2)
        tr = Right("00" & tr, 2)
        tc = Right("00" & tc, 2)
        
        
        USEITEMID = n2h(itemid, 2)
        senddata MakePacket("F4 44 0A 00 32 02 " & fr & " " & fc & " " & tr & " " & tc & " " & USEITEMID & " 0A A0")
        DoEvents
End Sub

'
'Public Sub SendPacket(str)
'        If str Mod 2 <> 0 Then
'            Exit Sub
'        End If
'        senddata MakePacket(str)
'        DoEvents
'End Sub
Public Sub RequestParty(tsid)
'F4 44 06 00 0D 01 88 57 0A 00"
     senddata MakePacket("F4 44 06 00 0D 01" & n2h(tsid, 4))
    DoEvents
End Sub


Private Sub packetBytes2StringAndDecode(data, ByVal bytesTotal As Long)
    If UBound(data) > 0 Then
    For i = 0 To bytesTotal - 1
        hx = Hex(DecodePacket(data(i)))
        StrBuffer = StrBuffer & Right("00" & hx, 2)
    Next
     Call SetBuffer(GetBuffer() & StrBuffer)
    StrBuffer = ""
x:
    header = Mid(GetBuffer(), 1, 4)
    If header = "F444" Then
        length = a2hex(Mid(dBuffer, 5, 4), 4)
        If (Len(dBuffer) - 8) >= (length * 2) Then
            bdata = Mid(dBuffer, 9, length * 2)
            sbuffer = header & n2h(length, 2) & bdata
            Call SetBuffer(Mid(dBuffer, 9 + Len(bdata)))
            Call ReceivePacket(sbuffer)
            GoTo x
        Else
            Exit Sub
        End If
    End If
    End If
End Sub
Private Function GetBuffer()
    GetBuffer = dBuffer
End Function
Private Sub SetBuffer(d)
    dBuffer = d
End Sub
Public Sub Loging(hdata)
    Dim fso As New Scripting.FileSystemObject
    Dim tso As Scripting.TextStream
    Set tso = fso.OpenTextFile(App.Path & "\log.txt", ForAppending, True, TristateUseDefault)
        tso.WriteLine Time & "-" & hdata
        tso.Close
    Set fso = Nothing
End Sub

Public Sub PrintP(ByRef packet() As Byte)
    Dim tstr As String
    Dim kstr As String
    Dim x As Long
    tstr = ""
    kstr = ""
    For x = 1 To UBound(packet)
       If packet(x) < 16 Then tstr = tstr + "0"
       tstr = tstr + Hex(packet(x)) + " "
       kstr = kstr + Chr(packet(x))
       If x Mod 16 = 0 Then
            tstr = tstr & vbTab & kstr & vbCrLf
            kstr = ""
        End If
    Next
       tstr = Left(tstr, Len(tstr) - 1)
       Open App.Path & "\unknow_packet.txt" For Append As #1
       If UBound(packet) Mod 16 = 0 Then
            Print #1, tstr
       Else
            Print #1, tstr & vbTab & kstr
       End If
       Print #1, "------------------"
       Close 1
End Sub

Public Sub PrintMail(mtype, uID, Msg, dtime)
        
        Open App.Path & "\Mail.txt" For Append As #1
        
        If mtype = 1 Then
            Print #1, "Mail type : กองทัพ"
        Else
            Print #1, "Mail type : สหายสนิท"
        End If
        Print #1, "Form : " & Form1.getPlayerName(uID) & " ID : TS" & uID
        Print #1, "Date/Time : " & dtime
        Print #1, Msg
        Print #1, "------------------"
        Close 1
End Sub

Private Sub ReceivePacket(hstr)
    Dim b() As Byte
    
    If (Len(hstr) Mod 2) <> 0 Then
        Exit Sub
    End If
    
    buff = hstr
    begin = 1
    pos = 1
    Do While pos > 0
        pos = InStr(begin, buff, "F444", vbBinaryCompare)
        begin = pos + 4
        cmdLength = Mid(buff, begin, 4)
        psize = a2hex(cmdLength, 4)
        hdata = Mid(buff, begin + 4, psize * 2)
        ReDim Preserve b(psize + 1)
        For i = 1 To psize
            b(i) = a2hex(Mid(hdata, (i * 2) - 1, 2), 2)
        Next
        hcmd = cInt1(b, 1)
        If hcmd = &H4 Then
            xxxx = hdata
        End If
        'Loging (hdata)
'        If Form1.chdebug.value = 1 Then
'            hcmd = cInt1(b, 1)
'            If hcmd <> &H2 Then
'               Debug.Print hdata
'            ElseIf hcmd <> &H3 Then
'               Debug.Print hdata
'            ElseIf hcmd <> &H4 Then
'               Debug.Print hdata
'            End If
'        End If
        
       ' Call PrintP(b)  ' เอาไว้ดู packet เกมส์
        Call parsePacket(b)
        begin = begin + 4 + psize
        pos = InStr(begin, buff, "F444", vbBinaryCompare)
    Loop

End Sub

Private Sub parsePacket(ByRef din() As Byte)
On Error GoTo ErrorHandler
    Dim pl As Character
    Dim hcmd
    Dim scmd
    Dim strTempx As String
    offset = 1
    hcmd = cInt1(din, offset)
    'MsgBox hcmd
    Select Case hcmd
        Case &H0
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H13
                    '// Login ซ้ำซ้อน
                  '  mvarobj.Close
                    RaiseEvent DuplicateLogin
                    Me.Disconect
                Case &H11
                    RaiseEvent PatchIncorrect
                    Me.Disconect
                Case &H47
                                        
                    Me.Disconect
            End Select
            '
        Case &H1
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1
                    uID = cInt4(din, 3)
                    If uID > 0 Then
                        If Me.Character.mapid = Me.ol.Item(uID).mapid Then
                            RaiseEvent PlayerLeaveMap(uID, Me.ol.Item(uID).mapid, Me.ol.Item(uID).x, Me.ol.Item(uID).y)
                        End If

                        RaiseEvent PlayerOffline(uID)
                        RaiseEvent ArmyListOffline(uID)
                        Me.LastPlayerLogoff = uID
                        If Me.ol.Exists(uID) Then
                            Me.ol.Remove (uID)
                        End If
                    End If
                    
                Case &H6
                    RaiseEvent LoginFail
                Case &H9
                    RaiseEvent ReadyForLogin
            End Select
        Case &H2
            'F4 44 0A 00 02 03 B5 89 01 00 67 69 67 73
            MessageType = cInt1(din, 2)
            If MessageType = 8 Then
                SenderId = 0
                MessageText = ""
            Else
                SenderId = cInt4(din, 3)
                MessageText = cStrN(din, 7, UBound(din) - 8)
                Me.LastMsgType = MessageType
                Me.LastMsgSenderId = SenderId
                Me.LastMessage = MessageText
            End If
            RaiseEvent ChatMessage(MessageType, MessageText, SenderId)
            DoEvents
'
        Case &H3 ' Player Or Someone Online : Recieve for some status
'       F4 44 31 00
'                   03
'                   88 57 0A 00
'                   01 00 00 01 32
'                   6A 01
'                   3F 07
'                   00 01 00 8F EA 76 1A 3E AE 7D 1A
'                   05 [F6 4E] [38 4D] [C3 36] [DE 52] [3C 5A]
'                   00 00 00 00 0B 00
'                   44 69 67 69 41 4C 4C
            uID = cInt4(din, 2)
            
'                If uid <> testid Then
'                    mvarobj.Close
'                    Exit Sub
'                End If
            
            'F4 44 31 00
            '03 88 57 0A 00 01 00 00
            '1F 2F BA 01 C7 01 00 01 00 8F EA 76 1A 3E AE 7D 1A 05 EF 50 2A 4D F4 2E
            
            If uID = Me.Character.uID Then
                    Me.Character.Ghost = cInt1(din, 7)
                    Me.Character.GodOfLuck = cInt1(din, 8)
                    Me.Character.mapid = cInt2(din, 9)
                    Me.Character.x = cInt2(din, 11)
                    Me.Character.y = cInt2(din, 13)
                    nitem = cInt1(din, 26)
                    begin = 27
                    For i = 1 To nitem
                        offset = begin + ((i - 1) * 2)
                        itemid = cInt2(din, offset)
                    Next
                    
                    begin = (offset + 2) + 7
                    If nitem = 0 Then begin = 34
                    pName = cStrN(din, begin, UBound(din) - begin - 1)
                    
                    Me.Character.PvPranking = cInt1(din, begin - 6)
'                    NewBorn = cInt1(din, begin - 2)
'                    HeroType = cInt1(din, begin - 1)
'                    Me.Character.NewBorn = NewBorn + HeroType
                   ' NewBorn = cInt1(din, begin - 3)
                   ' HeroType = cInt1(din, begin - 2)
                    Me.Character.NewBorn = NewBorn + HeroType
                    
                    Me.Character.charname = pName
                     Call ol.Add(Me.Character.uID, Me.Character)
                     '// alogin จะแสดง ตัวละคร
                     
                     '//  ตรงนี้ต้อง ตอบ ip กลับไปด้วย
                     ' F444 0B 00 2501 313237 2E 30 2E 30 2E 31
                    ip = mvarobj.RemoteHostIP
                    iplength = Len("" & ip)
                    
                    DataLength = Right("00" & Hex(iplength + 2), 2) & "00"
                    
                    hip = String2Hex(mvarobj.RemoteHostIP)
                    senddata MakePacket("F4 44" & DataLength & "2501" & hip)
                    DoEvents
                    Me.Character.isLicenseId = checkLicId(Me.Character.uID)
                    'MsgBox Me.Character.isLicenseId, vbRed
                    RaiseEvent Loginok
                    RaiseEvent PlayerOnline(Date)
                    
                    
            'Else
            '// Others
             'F4 44 39 00
             '03
             '08 C1 07 00
             '01 01 3E 00 00 56 51 96 07 7B 02 00 02 00 EF AF 7D 1A 3E AE 7D 1A
             '06 EF 50 FB 4C F3 2E DE 52 03 56 3C 5A 00 00 00 00 00 0B 00 AA D4 E0 C1 E2 A8 B4 EA D2 C2
             'F4 44 12 00 05 00 08 C1 07 00 EF 50 FB 4C F3 2E DE 52 03 56 3C 5A
             ElseIf uID > 0 Then
                Set pl = New Character
                  pl.uID = cInt4(din, 2)
                  pl.Element = cInt1(din, 7)
                  pl.level = cInt1(din, 8)
                    
                    pl.Ghost = cInt1(din, 9)
                    'pl.GodOfLuck = cInt1(din, 10)
                    pl.mapid = cInt2(din, 11)
                    pl.x = cInt2(din, 13)
                    pl.y = cInt2(din, 15)
                                
                    nitem = cInt1(din, 28)
                    begin = 29
                    For i = 1 To nitem
                        offset = begin + ((i - 1) * 2)
                        itemid = cInt2(din, offset)
                    Next
                    begin = (offset + 2) + 8
                    If nitem = 0 Then begin = 36
                    pName = cStrN(din, begin, UBound(din) - begin - 1)
                    
                pl.charname = pName
                
                    pl.PvPranking = cInt1(din, begin - 6)
                    'NewBorn = cInt1(din, begin - 2)
                    'HeroType = cInt1(din, begin - 1)
                    pl.NewBorn = NewBorn + HeroType
                    
                Me.LastPlayerOnline = pl.uID
                If Not ol.Exists(pl.uID) Then
                    Call ol.Add(pl.uID, pl)
                End If
                
                If Me.Character.mapid = pl.mapid Then
                    RaiseEvent PlayerAppearInMap(pl.uID, pl.x, pl.y)
                    If ((pl.uID < 10000) Or (pl.charname = "เทพบุตรฟ้าจุติ") Or (pl.charname = "เทพธิดาฟ้าจุติ") Or (pl.charname = "ธิดาฟ้าจุติ") Or (pl.charname = "ทารกฟ้าจุติ")) Then
                        Form1.AppendDisplay "Danger !! GM id = " & pl.uID & " name = " & pl.charname & " detected in your current map at = " & pl.mapid, vbRed
                        If (Form1.mnuGMInmap.Checked = True) And (pl.uID <> 101) Then
                            Form1.mnuEnableReconnect.Checked = False
                            Me.Disconect
                        End If
                    End If

'                Else
'                    RaiseEvent PlayerLeaveMap(cObj.uID, cObj.mapid, cObj.x, cObj.y)
                End If

                
                If ((pl.uID < 10000) Or (pl.charname = "เทพบุตรฟ้าจุติ") Or (pl.charname = "เทพธิดาฟ้าจุติ") Or (pl.charname = "ธิดาฟ้าจุติ") Or (pl.charname = "ทารกฟ้าจุติ")) Then
                    Form1.AppendDisplay "Warning !! GM id = " & pl.uID & " name = " & pl.charname & " detected at map = " & pl.mapid, vbRed
                    If (Form1.mnuGMInmap.Checked = True) And (pl.uID <> 101) Then
                        Form1.mnuEnableReconnect.Checked = False
                        Me.Disconect
                    End If
                End If
                
                RaiseEvent AppearOnlinePlayers(pl)
                RaiseEvent ArmyListOnline(pl.uID)
                
                Set pl = Nothing
            End If

        Case &H4 ' Current Online User on Server
                Set pl = New Character
'F444 38 00
'04
'01 53 05 00
'01
'03
'47
'00 00
'4C 51
'66 05
'A3 02
'00 02 00 A0 56 84 1A 50 8B EF 34
'06
'F5 4E
'0D 4B
'29 2B
'3B 53
'27 57
'32 5A
'00 00 00 00 00 0B 00
'41 71 75 61 72 69 6F 75 73
        
            
            uID = cInt4(din, 2)
            If uID > 0 Then
                pl.uID = uID
                pl.Element = cInt1(din, 7)
                pl.level = cInt1(din, 8)
                pl.mapid = cInt2(din, 11)
                                
                    nitem = cInt1(din, 28)
                    begin = 29
                    For i = 1 To nitem
                        offset = begin + ((i - 1) * 2)
                        itemid = cInt2(din, offset)
                    Next
                    
                    begin = (offset + 2) + 8
                    If nitem = 0 Then begin = 36
                    
                    pName = cStrN(din, begin, UBound(din) - begin - 1)
                
                pl.PvPranking = cInt1(din, begin - 6)
                NewBorn = cInt1(din, begin - 1)
                'HeroType = cInt1(din, begin - 1)
                pl.NewBorn = NewBorn + HeroType
                    
                pl.charname = pName
                Me.LastPlayerOnline = pl.uID
                If Not ol.Exists(pl.uID) Then
                    Call ol.Add(pl.uID, pl)
                End If
 
                RaiseEvent AppearCurrentOnlinePlayers(pl)
                RaiseEvent ArmyListOnline(pl.uID)
                If ((pl.uID < 10000) Or (pl.charname = "เทพบุตรฟ้าจุติ") Or (pl.charname = "เทพธิดาฟ้าจุติ") Or (pl.charname = "ธิดาฟ้าจุติ") Or (pl.charname = "ทารกฟ้าจุติ")) Then
                    Form1.AppendDisplay "Warning !! GM id = " & pl.uID & " name = " & pl.charname & " detected at map = " & pl.mapid, vbRed
                    If (Form1.mnuGMOnline.Checked = True) And (pl.uID <> 101) Then
                        Form1.mnuEnableReconnect.Checked = False
                        Me.Disconect
                    End If
                End If
            End If
                Set pl = Nothing
        
        Case &H5 ' About My player information
            
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H2
                'F4 44 08 00 05 02 AE F3 0B 00 34 5A
                
                Case &H3 ' My Current Status and more infor
                   
                     
                
                    Me.Character.Element = cInt1(din, 3)
                    Me.Character.HP = cInt2(din, 4)
                    Me.Character.SP = cInt2(din, 6)
                    'Me.Character.GodOfLuck = cInt1(din, 14)
                    'Me.Character.CharStatus
                    Me.Character.level = cInt1(din, 20)
                    Me.Character.Texp = cInt4(din, 21)
                    Me.Character.Hor = cInt4(din, 29)
                    Me.Character.MAXHP = cInt2(din, 33)
                    Me.Character.MAXSP = cInt2(din, 35)


                    
                    Me.Character.army_rate.Add 1, cInt2(din, 61)
                    Me.Character.army_rate.Add 2, cInt2(din, 63)
                    Me.Character.army_rate.Add 3, cInt2(din, 65)
                    Me.Character.army_rate.Add 4, cInt2(din, 67)
                    Me.Character.army_rate.Add 5, cInt2(din, 69)
                    
                    
                    offset = 110
                    Do While offset < UBound(din)
                        skillid = cInt2(din, offset)
                        Me.Character.Skills.Add Me.Character.Skills.Count, skillid
                        offset = offset + 3
                    Loop
                    RaiseEvent InitPlayerStatus
                    RaiseEvent onChangeStatus
                Case &H4
                    If WarpState = True Then
                                If Me.CurrentParty = 0 Then
                                    Me.SendEnd
                                End If

                        WarpState = False
                        RaiseEvent warpFinish
                    End If
                    
            End Select
'
        Case &H6
           'F4 44 02 00
           '06 02
           'F4 44 0B 00
           '06 01 32 3A 07 00 03 42 01 DB 01
           cmdtype = cInt1(din, 2)
           Select Case cmdtype
              Case &H1
                 '// Other walk
                  uID = cInt4(din, 3)
                  direction = cInt1(din, 7)
                  x = cInt2(din, 8)
                  y = cInt2(din, 10)
                  RaiseEvent PlayerWalk(uID, direction, x, y)
'                  Debug.Print "Other walk " & uid & " " & direction & " " & x & "," & y
               Case &H2
                  RequestToWork = True
           End Select
          
        Case &H8
                '//
'
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1
                    '// รับค่า status  จาก ตัวเอง
                    'F4 44 0C 00
                    '08 01 19 01 0A 03 00 00 00 00 00 00
                    vtype = cInt1(din, 3)
                    vData = cInt4(din, 5)
                    
                    
                    ' 19 - hp
                    ' 1A = sp
                    ' 24 - texp
                    Select Case vtype
                        Case &H19
                            Me.Character.HP = vData
                        Case &H1A
                            Me.Character.SP = vData
                        Case &H24
                            Me.Character.Texp = vData
                            RaiseEvent onUpdateTexp
                    End Select
                Case &H2
                    '// รับค่า status  จาก ขุนพล
                    'F4 44 0F 00
                    '08 02 04 01 00 19 01 3D 02 00 00 00 00 00 00
                        
                    vtype = cInt1(din, 6)
                    vorder = cInt1(din, 4)
                    vData = cInt4(din, 8)
                    Select Case vtype
                        Case &H19
                            Me.MyPartners(vorder - 1).HP = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.HP = Me.MyPartners(vorder - 1).HP
                            End If
                        Case &H1A
                            Me.MyPartners(vorder - 1).SP = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.SP = Me.MyPartners(vorder - 1).SP
                            End If
                        Case &H24
                            Me.MyPartners(vorder - 1).Texp = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.Texp = Me.MyPartners(vorder - 1).Texp
                                RaiseEvent onUpdateTexp
                            End If
                        Case &H23
                            Me.MyPartners(vorder - 1).level = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.level = Me.MyPartners(vorder - 1).level
                            End If
                        Case &H40
                            Me.MyPartners(vorder - 1).fai = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.fai = Me.MyPartners(vorder - 1).fai
                            End If
                        Case &H1B
                            Me.MyPartners(vorder - 1).CharStatus.INTX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.INTX = Me.MyPartners(vorder - 1).CharStatus.INTX
                            End If
                        Case &H1C
                            Me.MyPartners(vorder - 1).CharStatus.ATKX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.ATKX = Me.MyPartners(vorder - 1).CharStatus.ATKX
                            End If
                        Case &H1D
                            Me.MyPartners(vorder - 1).CharStatus.DEFX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.DEFX = Me.MyPartners(vorder - 1).CharStatus.DEFX
                            End If
                        Case &H1E
                            Me.MyPartners(vorder - 1).CharStatus.AGIX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.AGIX = Me.MyPartners(vorder - 1).CharStatus.AGIX
                            End If
                        Case &H1F
                            Me.MyPartners(vorder - 1).CharStatus.HPX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.HPX = Me.MyPartners(vorder - 1).CharStatus.HPX
                            End If
                        Case &H20
                            Me.MyPartners(vorder - 1).CharStatus.SPX = vData
                            If Me.MyPartners(vorder - 1).uID = Me.CurrentPartner.uID Then
                                Me.CurrentPartner.CharStatus.SPX = Me.MyPartners(vorder - 1).CharStatus.SPX
                            End If
                    End Select
                    Form1.Label1(1).ToolTipText = Me.CurrentPartner.charname & vbNewLine & _
                                                " LvL = " & Me.CurrentPartner.level & _
                                                " Status = " & Me.CurrentPartner.CharStatus.INTX & _
                                                " " & Me.CurrentPartner.CharStatus.ATKX & _
                                                " " & Me.CurrentPartner.CharStatus.DEFX & _
                                                " " & Me.CurrentPartner.CharStatus.HPX & _
                                                " " & Me.CurrentPartner.CharStatus.SPX & _
                                                " " & Me.CurrentPartner.CharStatus.AGIX & _
                                                " Fai = " & Me.CurrentPartner.fai

                Case &H3
                    '// รับค่า status  จาก คนอื่น (party)
                    'F4 44 10 00
                    '08 03 53 D7 08 00 1A 01 EC 00 00 00 00 00 00 00
                    pid = cInt4(din, 3)
                    vtype = cInt1(din, 7)
                    vData = cInt4(din, 9)
                     
                     Select Case vtype
                        Case &H19
                     '       Me.Character.HP = vData
                            RaiseEvent PartyUpdateHP(pid, vData)
                            
                        Case &H1A
                      '      Me.Character.SP = vData
                            RaiseEvent PartyUpdateSP(pid, vData)
                        Case &H24
                       '     Me.Character.Texp = vData
                        '    RaiseEvent onUpdateTexp
                     End Select
'
            End Select
            RaiseEvent onChangeStatus
        Case &HB
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H0
                '0B 00 A83A0400 0000
                    uID = cInt4(din, 3)
                    RaiseEvent FinishBattle(uID)
                Case &H1
'                'MsgBox UBound(din)
                         If UBound(din) = 6 Then '/// End combat
                            fightingstate = False
                            Set oNPCCombat = Nothing
                            
                            If godquestion = False Then
                                If Me.CurrentParty = 0 Then
                                    Me.SendEnd
                                    'senddata MakePacket("F444 0200 14 09 ")
                                End If
                                RaiseEvent onBattleStoped
                            Else
                                godquestion = True
                                If Me.CurrentParty = 0 Then
                                    Me.SendEnd
                                    'senddata MakePacket("F444 0200 14 09 ")
                                End If
                                'senddata MakePacket("F444 0200 14 09 ")
                                RaiseEvent onBattleStoped
                                RaiseEvent FinishAnswerFuckGod
                            End If
                        End If
                Case &H4
                        'F4 44 0A 00 0B 04 02 A8 3A 04 00 00 00 05
                        'F4 44 0A 00 0B 04 02 92 B1 0F 00 00 00 01
                        btype = cInt1(din, 3)
                        Select Case btype
                            Case &H2
                                uID = cInt4(din, 4)
                                RaiseEvent AppearAnotherCombat(uID)
                            
                        End Select
                Case &H5
'                       'F4 44 1A 00
                        '0B 05
                        '05 04
                        '2D 2B 00 00 01 00
                        '32 3A 07 00
                        '02 02
                        '3D 02
                        'D2 01
                        '3D 02
                        'D2 01
                        '5D
                        '03
                        Set npcscom = New NPCSCombat
                            npcscom.uID = cInt2(din, 5)
                            npcscom.ownerid = cInt4(din, 11)
                            npcscom.Row = cInt1(din, 15)
                            npcscom.Col = cInt1(din, 16)
                            npcscom.MAXHP = cInt2(din, 17)
                            npcscom.MAXSP = cInt2(din, 19)
                            npcscom.HP = cInt2(din, 21)
                            npcscom.SP = cInt2(din, 23)
                            npcscom.lv = cInt1(din, 25)
                            npcscom.elem = cInt1(din, 26)
                            
                            If npcscom.ownerid = Me.Character.uID Then
                              '  MsgBox Me.Character.uid
                            '// My partner
                                CurrentPartner.SP = npcscom.SP
                                CurrentPartner.HP = npcscom.HP
                                CurrentPartner.MAXHP = npcscom.MAXHP
                                CurrentPartner.MAXSP = npcscom.MAXSP
                                CurrentPartner.Element = npcscom.elem
                                CurrentPartner.level = npcscom.lv
                                CurrentPartner.Col = npcscom.Col
                                CurrentPartner.Row = npcscom.Row
                                CurrentPartner.uID = npcscom.uID
                                RaiseEvent onChangeStatus
                            Else
                            '// Party partner
                            End If
                            npctype = cInt2(din, 3)
                            If npctype = &H701 Then '/// NORMAL COMBAT NPC
                                Me.oNPCCombat.Add Me.oNPCCombat.Count, npcscom
                            ElseIf npctype = &H303 Then '/// PK NPC
                                Me.oNPCCombat.Add Me.oNPCCombat.Count, npcscom
                            End If
                    
                    Set npcscom = Nothing
                Case &HFA    '// Start Battle
                        Set oNPCCombat = New Dictionary
                        godquestion = False
                        tep = 0

'F4 44 AC 00

'    0B FA 6D 07
'    05 02 DC E8 05 00 00 00 00 00 00 00 03 00 D4 01 B9 00 D4 01 B9 00 64 04
'    01 02 32 3A 07 00 00 00 00 00 00 00 03 02 0D 03 C0 00 0D 03 C0 00 64 03
'    01 04 2D 2B 00 00 00 00 32 3A 07 00 02 02 40 02 D4 01 40 02 8A 01 5E 03
'    64 02 53 D7 08 00 00 00 00 00 00 00 03 01 19 03 10 01 19 03 EE 00 59 02
'    64 04 2D 2B 00 00 00 00 53 D7 08 00 02 01 4D 02 88 01 4D 02 14 01 47 03
'    64 02 88 57 0A 00 00 00 00 00 00 00 03 03 80 01 8B 00 80 01 8B 00 4F 04
'    64 04 2D 2B 00 00 00 00 88 57 0A 00 02 03 16 02 3D 01 16 02 E7 00 32 03
                offset = 5
                    Do While offset < UBound(din)
                            npctype = cInt2(din, offset)
                        Set mvarNPCSCombat = New NPCSCombat
                        
                            mvarNPCSCombat.uID = cInt4(din, offset + 2)
                            mvarNPCSCombat.Row = cInt1(din, offset + 12)
                            mvarNPCSCombat.Col = cInt1(din, offset + 13)
                            mvarNPCSCombat.MAXHP = cInt2(din, offset + 14)
                            mvarNPCSCombat.MAXSP = cInt2(din, offset + 16)
                            mvarNPCSCombat.HP = cInt2(din, offset + 18)
                            mvarNPCSCombat.SP = cInt2(din, offset + 20)
                            mvarNPCSCombat.lv = cInt1(din, offset + 22)
                            mvarNPCSCombat.elem = cInt1(din, offset + 23)
                      '  oNPCCombat.Add oNPCCombat.Count, mvarNPCSCombat
                        'MsgBox mvarNPCSCombat.uid & " " & offset
                        
                        If mvarNPCSCombat.uID = Me.Character.uID Then
                            Me.Character.Row = mvarNPCSCombat.Row
                            Me.Character.Col = mvarNPCSCombat.Col
                            Me.Character.level = mvarNPCSCombat.lv
'                            Me.Character.MAXHP = mvarNPCSCombat.MAXHP
'                            Me.Character.MAXSP = mvarNPCSCombat.MAXSP
'                            Me.Character.HP = mvarNPCSCombat.HP
'                            Me.Character.SP = mvarNPCSCombat.SP
                            
                        End If
                        If npctype = &H764 Then
                            Me.oNPCCombat.Add Me.oNPCCombat.Count, mvarNPCSCombat
                        End If
                        
                        
                        offset = offset + 24
                        Set mvarNPCSCombat = Nothing
                    Loop
'                    RaiseEvent onChangeStatus
                     RaiseEvent onBattleStarted
 
            End Select
''
''
        Case &HC
            'F4 44 0D 00 0C 92 B1 0F 00 08 32 5E 01 9A 06 02 00
            uID = cInt4(din, 2)
            mapid = cInt2(din, 6)
            x = cInt2(din, 8)
            y = cInt2(din, 10)
            warpid = cInt1(din, 12)
            If uID = Me.CurrentParty Then
                Me.LastWarpId = warpid
                Me.LastLeaderWarpId = warpid
            End If
            If uID = Me.Character.uID Then
                Me.Character.x = x
                Me.Character.y = y
                Me.Character.mapid = mapid
                If warpid = 0 Then
                    warpid = Me.LastLeaderWarpId
                End If
                WarpState = True
                'F4 44 02 00 0C 01
                senddata MakePacket("F4 44 02 00 0C 01")
                RaiseEvent onWarp(uID, mapid, warpid)
                Me.lastMapID = mapid
            Else
                Dim cObj As Character
                Set cObj = ol.Item(uID)
                    cObj.mapid = mapid
                    cObj.x = x
                    cObj.y = y
                Set ol.Item(uID) = cObj
                If mapid = Me.Character.mapid Then
                    RaiseEvent PlayerAppearInMap(cObj.uID, cObj.x, cObj.y)
                Else
                    RaiseEvent PlayerLeaveMap(cObj.uID, cObj.mapid, cObj.x, cObj.y)
                End If

            End If
            
        Case &HD
        '   F4 44 0B 00 0D 06 A8 3A 04 00 01 88 57 0A 00
        '   F4 44 06 00 0D 09 32 3A 07 00
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1
                    requestid = cInt4(din, 3)
                    Me.LastRequestPartyId = requestid
                    'accept
                    'F4 44 07 00 0D 03 01 88 57 0A 00
                    RaiseEvent WaitngForAcceptParty(requestid)
                    'senddata MakePacket("F4 44 07 00 0D 03 01" & n2h(requestid, 4))
                Case &H3
                    '0D03 01 A83A0400
                    '0D03 03 A83A0400
                    '0D03 02 A83A0400
                    f = cInt1(din, 3)
                    Select Case f
                        Case &H1
                            uID = cInt4(din, 4)
                            partyLeader = False
                            RaiseEvent RequestPartyAcceptFrom(uID)
                        Case &H2
                            uID = cInt4(din, 4)
                            RaiseEvent RequestPartyCancel(uID)
                        Case &H3
                            uID = cInt4(din, 4)
                            RaiseEvent RequestPartyFalse(uID)
                    End Select
                Case &H4
                    '0D 04 A83A0400
                    uID = cInt4(din, 3)
                    RaiseEvent PartyStop(uID)
                Case &H5
                'F4 44 0A 00 0D 05 A8 3A 04 00 92 B1 0F 00
                    LeaderID = cInt4(din, 3)
                    partyid = cInt4(din, 7)
                    If LeaderID = Me.Character.uID Then
                        RaiseEvent doAcceptParty(partyid)
                    End If
                Case &H9
                    requestid = cInt4(din, 3)
                    Me.LastRequestPartyId = requestid
                Case &H7
                'F4 44 06 00 0D 0B 9F F8 1F 00
                    uID = cInt4(din, 3)
                    RaiseEvent onSetsena(uID)
            End Select
        Case &HE
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1    '// Friend Mail
                    ktype = 2
                    kuid = cInt4(din, 3)
                    kmsg = cStrN(din, 15, UBound(din) - 16)
                    kdate = CDate(Hex2Double(Mid(bytes2hexStr(din), 13, 16)))
                    Call PrintMail(ktype, kuid, kmsg, kdate)
            End Select
        Case &HF
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H8    '// PartnerList
'F4 44 ED 00
'0F 08
'01 2D 2B 23 C9 0C 00 2E E0 01 0A 00 7B 00 14 00 36 00 46 00 28 00 26 00 00 22 01 00 00 05 A1 D8 C2 E1 A1 0A 0A 0A B9 4F 07 00 00 00 CF 4B AE 00 00 00 5A 46 7E 00 00 00 A8 53 03 00 00 00 88 57 7D 00 00 00 D9 59 00 00 00 00 00
'02 0D 56 0A 00 00 00 01 81 00 43 00 05 00 04 00 05 00 04 00 0C 00 03 00 00 3B 01 00 00 0B BB D2 E2 B5 E9 E0 C2 D2 E0 BB E9 01 00 00 00 00 00 00 00 00 39 4A 9C 00 00 00 11 27 19 00 00 00 00 00 00 00 00 00 81 57 9C 00 00 00 B6 B3 00 00 00 00 00
'03 5A 46 06 00 00 00 01 C9 00 61 00 05 00 1C 00 28 00 35 00 1E 00 12 00 00 3C 01 00 00 09 C1 E9 D2 B5 D5 EB CB C5 D9 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 11 27 7F 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00

'F4 44 49 00
'0F 08
'01 19 37 95 01 00 00 06 61 00 58 00 04 00 01 00 03 00 02 00 02 00 07 00 00 41 01 05 00 07 53 6F 6E 67 52 65 6E 01 00 00 00 00 00 00 00 39 4A 00 00 00 11 27 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00

                    offset = 3
                    n = 0
                    Do While offset < UBound(din)
                    Set partner = New Character

                        partner.Order = cInt1(din, offset)
                        partner.uID = cInt2(din, offset + 1)
                        partner.Texp = cInt4(din, offset + 3)
                        partner.level = cInt1(din, offset + 7)

                        partner.HP = cInt2(din, offset + 8)
                        partner.SP = cInt2(din, offset + 10)
                        partner.CharStatus.INTX = cInt2(din, offset + 12)
                        partner.CharStatus.ATKX = cInt2(din, offset + 14)
                        partner.CharStatus.DEFX = cInt2(din, offset + 16)
                        partner.CharStatus.AGIX = cInt2(din, offset + 18)
                        partner.CharStatus.HPX = cInt2(din, offset + 20)
                        partner.CharStatus.SPX = cInt2(din, offset + 22)
                        partner.fai = cInt1(din, offset + 25)
                        namelength = cInt1(din, offset + 29)

                        partner.charname = cStrN(din, offset + 30, namelength - 1)
                        partner.IsSleep = True
                        
                        'offset = offset + 30 + namelength + 69 ' ของ ep9
                        offset = offset + 30 + namelength + 69
                       'MsgBox offset
                        Call MyPartners.Add(MyPartners.Count, partner)
                        RaiseEvent RecvPartnerLists(partner)
                    Set partner = Nothing
                    Loop
                  '  RaiseEvent onChangeStatus
            End Select
'
'' รับ
''F4 44 60 01 27 1F 03 00 00 00 08 05 00 00 00 08 07 00 00 00 5B 17 00 00 00 4C 20 00 00 00 21 0D 01 00 00
''36 6C 00 00 00 A8 2C 00 00 00 2A 30 00 00 00 9F 38 00 00 00 15 46 00 00 00 21 62 00 00 00 75 63 00 00 00
''09 6B 00 00 00 02 72 00 00 00 04 7A 00 00 00 19 BA 00 00 00 4B C2 00 00 00 49 DF 00 00 00 55 D0 00 00 00
''74 E0 00 00 00 12 4C 06 00 00 04 F4 00 00 00 04 FD 00 00 00 1C 41 01 00 00 92 43 01 00 00 05 7A 01 00 00
''21 69 01 00 00 41 BB 01 00 00 1F E7 01 00 00 08 9F 02 00 00 11 E2 02 00 00 42 DE 02 00 00 2F F9 02 00 00
''06 5D 03 00 00 4F 45 03 00 00 17 6D 04 00 00 09 70 04 00 00 17 29 05 00 00 05 E4 04 00 00 1A 53 05 00 00
''01 91 05 00 00 03 5F 05 00 00 10 F1 06 00 00 6C 77 05 00 00 06 36 07 00 00 0C 51 07 00 00 1F D8 08 00 00
''23 EB 08 00 00 19 F4 0B 00 00 13 5B 0C 00 00 53 61 0E 00 00 0A FF 0D 00 00 0F 60 0F 00 00 54 BE 0F 00 00
''5C 04 11 00 00 1D DD 10 00 00 56 76 12 00 00 0B F3 11 00 00 27 68 11 00 00 0A 8F 12 00 00 01 C0 12 00 00
''01 D2 11 00 00 0F E8 12 00 00 04 EF 12 00 00 07 90 12 00 00 06 60 12 00 00 1C 67 12 00 00 05 81 12 00 00
''11 F8 12 00 00 01
'
'' ตอบ
''F4 44 1A 00 27 13 51 070000EB08000068110000E812000081120000F8120000
'

        Case &H10
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1
                    godquestion = True
                    'F4 44 0F 00
                    '10
                    '01 05 38 2D 35 3D 3F
                    '01 37 01 33  01 39 38
                    questLength = cInt1(din, 3)
                    questStr = cStrN(din, 4, questLength - 1)
                    
                    offset = 4 + questLength
                    ans1length = cInt1(din, offset)
                    ans1Str = cStrN(din, offset + 1, ans1length - 1)
                    
                    offset = offset + 1 + ans1length
                    ans2length = cInt1(din, offset)
                    ans2Str = cStrN(din, offset + 1, ans2length - 1)
'                    MsgBox ans2Str
                    
                    offset = offset + 1 + ans2length
                    ans3length = cInt1(din, offset)
                    ans3Str = cStrN(din, offset + 1, ans3length - 1)
                    'MsgBox ans3Str
                    
                    offset = offset + 1 + ans3length
                   ' MsgBox UBound(din)
'                    ans4length = cInt1(din, offset)
                    ans4Str = cStrN(din, offset, (UBound(din) - 1) - offset)
                   ' MsgBox ans4Str
                    Set Me.LastAnswers = New Dictionary
                    Me.LastQuestion = questStr
'                    Me.LastAnswers.Add 1, ans1Str
'                    Me.LastAnswers.Add 2, ans2Str
'                    Me.LastAnswers.Add 3, ans3Str
'                    Me.LastAnswers.Add 4, ans4Str
                    Me.LastAnswers.Add ans1Str, 1
                    Me.LastAnswers.Add ans2Str, 2
                    Me.LastAnswers.Add ans3Str, 3
                    Me.LastAnswers.Add ans4Str, 4
                    
                   RaiseEvent RecvQuestion
                Case &H2
                    ans_result = cInt1(din, 3)
                    Select Case ans_result
                        Case &H1
                            RaiseEvent onAnswerRight(Me.LastQuestion, Me.LastResponseAnswer)
                        Case &H2
                            RaiseEvent onAnswerWrong(Me.LastQuestion, Me.LastResponseAnswer)
                    End Select
            End Select
        Case &H13
'        '13 04 052B 0000
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1    '// กดออกรบ
'            'F4 44 06 00 13 01 F3 55 00 00
'            'F4 44 06 00 13 01 05 2B 00 00
                    selectpartnerid = cInt2(din, 3)
                    For i = 0 To Me.MyPartners.Count - 1
                        Me.MyPartners.Item(i).IsSleep = True
                        If Me.MyPartners.Item(i).uID = selectpartnerid Then
                            Me.MyPartners.Item(i).IsSleep = False
                            Set CurrentPartner = Me.MyPartners.Item(i)
                        End If
                    Next
'
                    RaiseEvent DoSelectPartner(selectpartnerid)
                    RaiseEvent onChangeStatus
'
                Case &H2    '// กดพักผ่อน
        '            'F4 44 02 00 13 02
                        Me.CurrentPartner.IsSleep = True
                    RaiseEvent DoUnSelectPartner
                Case &H4    '// ตัวที่ออกรบอยู่
                    selectpartnerid = cInt2(din, 3)
                    'MsgBox selectpartnerid
                    For i = 0 To Me.MyPartners.Count - 1
                    'MsgBox Me.MyPartners.Item(i).uid
                        Me.MyPartners.Item(i).IsSleep = True
                        If Me.MyPartners.Item(i).uID = selectpartnerid Then
                            Me.MyPartners.Item(i).IsSleep = False
                            Set CurrentPartner = Me.MyPartners.Item(i)
                        End If
                    Next
'
                    RaiseEvent DoSelectPartner(selectpartnerid)
                    RaiseEvent onChangeStatus
                End Select
        Case &H14
            
           ' MsgBox c
            cmd = cInt1(din, 2)
'F4 44 02 00 06 02
'F4 44 11 00 14 01 00 00 00 01 06 03 01 00 00 00 00 00 00 02 00
                Select Case cmd
                    Case &H1
'                    'F4 44 11 00 14 01 00 00 00 01 01 03 02 00 00 00 00 00 00 E1 37
                        dtype = cInt1(din, 7)
                        Me.lastNPCtalk = cInt1(din, 6)
                        mvarDialogId = cInt2(din, 16)
                        Select Case dtype
                            Case &H0
                                SendEnd
                            Case &H1
                                RaiseEvent NpcDialog(mvarDialogId)
                            Case &H4
                                mvarDialogId = cInt2(din, 12)
                                RaiseEvent NpcWalkThenDialog(mvarDialogId)
                                senddata MakePacket("F4 44 02 00 0A 01")
                                DoEvents
                                senddata MakePacket("F4 44 12 00 27 13 99 01 00 00 E2 01 00 00 DF 01 00 00 EC 01 00 00")
                            Case &H5
                                mvarDialogId = cInt2(din, 12)
                                RaiseEvent CombatSceneDialog(mvarDialogId)
                            Case &H6
                                RaiseEvent NpcDialogMenu(mvarDialogId)
                                
                            
                            Case Else
                                SendEnd
                            
                        End Select
'
                    Case &H8
                        '/// warp response
                        If requesttowarp = True Then
                                ' Me.LastWarpId = warpid
                                requesttowarp = False
                               ' warpIdhex = n2h(warpid, 2)
                                senddata MakePacket("F4 44 04 00 14 08 " & n2h(Me.LastWarpId, 2))
                                DoEvents
                        End If
                        'If RequestToWork = True Then
                        '    RequestToWork = False
                        '    senddata MakePacket("F4 44 03 00 20 02 23")
                        '    DoEvents
                        'End If
                    Case &HD
                        SendEnd
                    Case &H10
                        SendEnd
                    Case &HB
                     '   SendEnd
                         RaiseEvent on140B
                    
                End Select
'
        Case &H16
        'F4 44 08 00 16 02 01 00 CB 02 8D 03
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H2
                    npcmapid = cInt1(din, 3)
                    npconmapx = cInt2(din, 5)
                    npconmapy = cInt2(din, 7)
                    RaiseEvent onNPCAppear(npcmapid, npconmapx, npconmapy)
            End Select
        
        Case &H17
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H5
'F4 44 2C 00 17 05
'01 0C 7D 04 00 00 00
'02 79 69 01 00 00 00 03 11 27 01 00 00 00 04 91 65 01 00 00 00 05 39 4A 01 00 00 00 06 43 A8 02 00 00 00

                
'F4 44 C2 00
'17 05
'01 CB B3 1D 00 00 00 00
'02 C0 B3 2D 00 00 00 00 03 E4 36 01 00 00 00 00 04 D4 B3 04 00 00 00 00 05 D9 B3 02 00 00 00 00 06 89 50 01 00 00 00 00 07 30 5A 01 00 00 00 00 08 30 5A 01 00 00 00 00 09 C4 5D 02 00 00 00 00 0A 90 69 2B 00 00 00 00 0B D1 B3 11 00 00 00 00 0C 04 7D 1B 00 00 00 00 0D B3 65 04 00 00 00 00 0E F5 65 12 00 00 00 00 0F BD B3 1C 00 00 00 00 10 51 C3 04 00 00 00 00 11 F2 65 0A 00 00 00 00 12 AE 65 06 00 00 00 00 13 A0 65 1B 00 00 00 00 14 03 66 25 00 00 00 00 16 E0 65 10 00 00 00 00 17 8E 69 06 00 00 00 00 18 DE 65 03 00 00 00 00 19 7B 69 03 00 00 00 00

                    offset = 3
                    Do While offset < UBound(din)
                        slotno = cInt1(din, offset)
                        
                        itemid = cInt2(din, offset + 1)
                        If Not ditems.Exists(itemid) Then
                            itemid = 99999
                        End If
                        num = cInt1(din, offset + 3)
                        fine = cInt1(din, offset + 4)
'                        un = a2hex(Mid(hexstring, offset + 8, 8), 8)
                        'Debug.Print slotno & "," & itemid
                        
                            Me.MyItems.Item(slotno).slot = slotno
                            Me.MyItems.Item(slotno).itemid = itemid
                            Me.MyItems.Item(slotno).num = num
                            Me.MyItems.Item(slotno).fine = fine
                        
                        offset = offset + 12
                    Loop
                    RaiseEvent InitInventoryList
                Case &H6
'                    '17 06 D8B3 01 00000000
'                '17 06 DB84 01 0000000000
                        itemid = cInt2(din, 3)
                        If Not ditems.Exists(itemid) Then
                            itemid = 99999
                        End If
                        num = cInt1(din, 5)
                        'MsgBox itemid & " " & num
                        haveitem = False
                        RaiseEvent RecvDropItems(itemid, num)
                        For i = 1 To 25
                            If itemid = 23086 Or _
                                itemid = 23087 Or _
                                itemid = 23088 Or _
                                itemid = 23089 Or _
                                itemid = 23028 Or _
                                itemid = 23029 Or _
                                itemid = 23023 Or _
                                itemid = 23085 Or _
                                itemid = 20737 _
                            Then
                                Exit For
                            End If
                            If Me.MyItems.Item(i).itemid = itemid Then
                                haveitem = True
                                newnum = Me.MyItems.Item(i).num + num
                                If newnum > 50 Then
                                    Me.MyItems.Item(i).num = 50
                                    num = newnum - 50
                                Else
                                    Me.MyItems.Item(i).num = newnum
                                    num = 0
                                    Exit For
                                End If
                            End If
                        Next
                        haveslot = False
                        If haveitem = False Or num > 0 Then
                            For i = 1 To 25
                                If Me.MyItems.Item(i).num = 0 Then
                                    haveslot = True
                                    Me.MyItems.Item(i).slot = i
                                    Me.MyItems.Item(i).itemid = itemid
                                    Me.MyItems.Item(i).num = num
                                    Exit For
                                 End If
                            Next
                            If haveslot = False Then
                                RaiseEvent doNotEnoughSlot(itemid, num)
                            End If
                        End If
                        
                        
                        RaiseEvent InventoryChange
                Case &H8
                '17 08 09 599C 04 000000
                        slot = cInt1(din, 3)
                        itemid = cInt2(din, 4)
                        If Not ditems.Exists(itemid) Then
                            itemid = 99999
                        End If
                        num = cInt1(din, 6)
                            Me.MyItems.Item(slot).slot = slot
                            Me.MyItems.Item(slot).itemid = itemid
                            Me.MyItems.Item(slot).num = num
                        RaiseEvent InventoryChange
                Case &H9
                '1709 16 01
                    slot = cInt1(din, 3)
                    num = cInt1(din, 4)
                    Me.MyItems.Item(slot).num = Me.MyItems.Item(slot).num - num
                    If Me.MyItems.Item(slot).num <= 0 Then
                        Me.MyItems.Item(slot).itemid = 0
                    End If
                    
                    RaiseEvent InventoryChange
                Case &HA
                'F4 44 05 00 17 0A 05 01 04
                    FromSlot = cInt1(din, 3)
                    ToSlot = cInt1(din, 5)
                    num = cInt1(din, 4)
                    If Me.MyItems.Item(ToSlot).itemid = 0 Then
                        Me.MyItems.Item(ToSlot).itemid = Me.MyItems.Item(FromSlot).itemid
                    End If
                        
                        Me.MyItems.Item(FromSlot).num = Me.MyItems.Item(FromSlot).num - num
                        Me.MyItems.Item(ToSlot).num = Me.MyItems.Item(ToSlot).num + num
                        If Me.MyItems.Item(FromSlot).num <= 0 Then
                            Me.MyItems.Item(FromSlot).itemid = 0
                        End If
                    
                    RaiseEvent InventoryChange
                Case &HF
                
                Case &H11
                'F4 44 03 00 17 11 19
                    slot = cInt1(din, 3)
                    RaiseEvent onEquipment(slot)
                
                Case &H22
                'F4 44 0A 00 17 22 01 CB B3 02 0A 00 00 00
                ' ขายออก
                    itemid = cInt2(din, 4)
                    num = cInt1(din, 6)
                    money = cInt4(din, 7)
                    RaiseEvent onSales(itemid, num, money)
                    
                Case &H2F
                'รับข้อมูลปาเป้ตอนเข้าเกม
                    offset = 3
                    Do While offset < UBound(din)
                        slotno = cInt1(din, offset)
                        itemid = cInt2(din, offset + 1)
                        If Not ditems.Exists(itemid) Then
                            itemid = 99999
                        End If

                        num = cInt1(din, offset + 3)
                        fine = cInt1(din, offset + 4)
'                        un = a2hex(Mid(hexstring, offset + 8, 8), 8)
                        'Debug.Print slotno & "," & itemid
                        
                            Me.BackPack.Item(slotno).slot = slotno
                            Me.BackPack.Item(slotno).itemid = itemid
                            Me.BackPack.Item(slotno).num = num
                            Me.BackPack.Item(slotno).fine = fine
                        
                        offset = offset + 12
                    Loop
                    RaiseEvent InitBackPackList
                Case &H30
                'ได้รับของเข้าปาเป้
                    offset = 3
                        'slotno = cInt1(din, offset)
                        itemid = cInt2(din, offset)
                        If Not ditems.Exists(itemid) Then
                            itemid = 99999
                        End If
                        num = cInt1(din, offset + 2)
                        'fine = cInt1(din, offset + 4)
'                        un = a2hex(Mid(hexstring, offset + 8, 8), 8)
                        'Debug.Print slotno & "," & itemid
                        haveitem = False
                        RaiseEvent RecvBackPackItems(itemid, num)
                        For i = 1 To 25
                            If itemid = 23086 Or _
                                itemid = 23087 Or _
                                itemid = 23088 Or _
                                itemid = 23089 Or _
                                itemid = 23028 Or _
                                itemid = 23029 Or _
                                itemid = 23023 Or _
                                itemid = 23085 _
                            Then
                                Exit For
                            End If
                            If Me.BackPack.Item(i).itemid = itemid Then
                                haveitem = True
                                newnum = Me.BackPack.Item(i).num + num
                                If newnum > 50 Then
                                    Me.BackPack.Item(i).num = 50
                                    num = newnum - 50
                                Else
                                    Me.BackPack.Item(i).num = newnum
                                    num = 0
                                    Exit For
                                End If
                            End If
                        Next
                        haveslot = False
                        If haveitem = False Or num > 0 Then
                            For i = 1 To 25
                                If Me.BackPack.Item(i).num = 0 Then
                                    haveslot = True
                                    Me.BackPack.Item(i).slot = i
                                    Me.BackPack.Item(i).itemid = itemid
                                    Me.BackPack.Item(i).num = num
                                    Exit For
                                 End If
                            Next
                            If haveslot = False Then
                                RaiseEvent doNotEnoughBackPackSlot(itemid, num)
                            End If
                        End If
                        
                            'Me.BackPack.Item(slotno).slot = slotno
                            'Me.BackPack.Item(slotno).itemid = itemid
                            'Me.BackPack.Item(slotno).num = num
                            'Me.BackPack.Item(slotno).fine = fine
                       
                    RaiseEvent BackPackChange
                Case &H31
                    'ของหายจากเป้
                    offset = 3
                    slotno = cInt1(din, offset)
                            Me.BackPack.Item(slotno).slot = slotno
                            Me.BackPack.Item(slotno).itemid = 0
                            Me.BackPack.Item(slotno).num = 0
                            Me.BackPack.Item(slotno).fine = 0
                    RaiseEvent BackPackChange
            End Select
        Case &H19
        'F4 44 0D 00 19 03 00 00 00 00 D1 B3 01 00 00 00 00
            Deal = cInt1(din, 2)
            Debug.Print Deal
        
        'F4 44 0A 00 19 17 01 32 3A 07 00 D1 B3 07
        
            ctype = cInt2(din, 2)
            Select Case ctype
                Case &H117
                '/// รับของ จาก
                    uID = cInt4(din, 4)
                    itemid = cInt2(din, 8)
                    num = cInt1(din, 10)
                    RaiseEvent RecvItemFrom(uID, itemid, num)
                Case &H217
                '/// ส่งของให้
                    uID = cInt4(din, 4)
                    itemid = cInt2(din, 8)
                    num = cInt1(din, 10)
                    RaiseEvent SendItemSuccess(uID, itemid, num)
            End Select
        'F4 44 0A 00 19 17 02 32 3A 07 00 D1 B3 07
    
        
'
'F4 44 04 00
'35 05 02 02
        Case &H1A
          '  F4 44 06 00 1A 01 08 00 00 00
          '  F4 44 0A 00 1A 04 0F 00 00 00 00 00 00 00
            ctype = cInt1(din, 2)
            Select Case ctype
                Case &H1
                    money = cInt4(din, 3)
                    Me.Character.gold = Me.Character.gold + money
                    RaiseEvent RecvMoney(money)
                Case &H2
                    money = cInt4(din, 3)
                    Me.Character.gold = Me.Character.gold - money
                    RaiseEvent RecvMoney(money)
                Case &H4
                    money = cInt4(din, 3)
                    Me.Character.gold = Me.Character.gold + money
                    RaiseEvent RecvMoney(money)
            End Select
        Case &H1F
            ' นอนเตี๊ยม
            'F4 44 07 00 1F 02 28 00 00 00 01
            ctype = cInt1(din, 2)
            Select Case ctype
                Case &H2
                    price = cInt1(din, 3)
                    RaiseEvent onRequestSleep(price)
                'F4 44 02 00 14 09
            End Select
            
        Case &H23
            'Check Airtime
            '23 04 00 00 00 00 53 03 00 00 46 F2 33 11 1D E4 E2 40
            ctype = cInt1(din, 2)
            Select Case ctype
                Case &H4
                    timemin = cInt2(din, 7)
                    Form1.AppendChat "สามารถเล่นได้อีก > " & timemin & " นาที", vbBlue
                    Form1.AppendDisplay "สามารถเล่นได้อีก > " & timemin & " นาที", vbBlue
                    'MsgBox "hex = " & bytes2hexStr(din) & vbCrLf & Mid(bytes2hexStr(din), 21, 16)
                    DateFinal = CDate(Hex2Double(Mid(bytes2hexStr(din), 21, 16)))
                    'MsgBox "time = " & timemin & vbCrLf & "date = " & datefinal
                    Form1.AppendChat "สามารถเล่นได้ถึง > " & DateFinal, vbBlue
                    Form1.AppendDisplay "สามารถเล่นได้ถึง > " & DateFinal, vbBlue
                    CurrentDate = Date
                    If Now > DateFinal Then
                        If timemin > 0 Then
                            Form1.AppendChat "ขณะนี้เหลือแต้มเฉพาะระบบ ชม. แล้วนะจ๊ะ", &H10AA10
                            Form1.AppendDisplay "ขณะนี้เหลือแต้มเฉพาะระบบ ชม. แล้วนะจ๊ะ", &H10AA10
                        Else
                            Form1.AppendChat "หมดเวลาแล้วจ้า อย่าลืมเติมเงินด้วยนะ ^^V", vbRed
                            Form1.AppendDisplay "หมดเวลาแล้วจ้า อย่าลืมเติมเงินด้วยนะ ^^V", vbRed
                        End If
                    End If
                'F4 44 02 00 14 09
            End Select
        Case &H27
            ctype = cInt1(din, 2)
            Select Case ctype
                Case &H2
                offset = 3
                ArmyNameLen = cInt1(din, offset)
                myArmy.AName = Space$(ArmyNameLen)
                myArmy.AName = cStrN(din, offset + 1, ArmyNameLen - 1)
                'Form1.TabStrip1.Tabs.Item(3).Caption = "Army - " & myArmy.AName
                offset = offset + ArmyNameLen + 1
                myArmy.ASubCount = cInt1(din, offset)
                myArmy.AMemCount = cInt1(din, offset + 1)
                
                offset = offset + 2
                myArmy.ALeader.id = cInt4(din, offset)
                NameLen = cInt1(din, offset + 4)
                myArmy.ALeader.Name = cStrN(din, offset + 5, NameLen - 1)
                myArmy.ALeader.Lvl = cInt1(din, offset + NameLen + 5)
                myArmy.ALeader.ele = cInt1(din, offset + NameLen + 6)
                myArmy.ALeader.Reborn = cInt1(din, offset + NameLen + 17)
                myArmy.ALeader.Con = cInt4(din, offset + NameLen + 18)
                offset = offset + NameLen + 27
                For intSubCount = 1 To myArmy.ASubCount
                    myArmy.ASubLeader(intSubCount).id = cInt4(din, offset)
                    NameLen = cInt1(din, offset + 4)
                    myArmy.ASubLeader(intSubCount).Name = cStrN(din, offset + 5, NameLen - 1)
                    myArmy.ASubLeader(intSubCount).Lvl = cInt1(din, offset + NameLen + 5)
                    myArmy.ASubLeader(intSubCount).ele = cInt1(din, offset + NameLen + 6)
                    myArmy.ASubLeader(intSubCount).Reborn = cInt1(din, offset + NameLen + 17)
                    myArmy.ASubLeader(intSubCount).Con = cInt4(din, offset + NameLen + 18)
                    offset = offset + NameLen + 27
                Next intSubCount
                For intSubCount = 1 To myArmy.AMemCount
                    myArmy.AMember(intSubCount).id = cInt4(din, offset)
                    NameLen = cInt1(din, offset + 4)
                    myArmy.AMember(intSubCount).Name = cStrN(din, offset + 5, NameLen - 1)
                    myArmy.AMember(intSubCount).Lvl = cInt1(din, offset + NameLen + 5)
                    myArmy.AMember(intSubCount).ele = cInt1(din, offset + NameLen + 6)
                    myArmy.AMember(intSubCount).Reborn = cInt1(din, offset + NameLen + 17)
                    myArmy.AMember(intSubCount).Con = cInt4(din, offset + NameLen + 18)
                    offset = offset + NameLen + 27
                Next intSubCount
                
                RaiseEvent InitArmyList

'F4 44 C2 00 'backpack
'17 05
'01 CB B3 1D 00 00 00 00
'02 C0 B3 2D 00 00 00 00 03 E4 36 01 00 00 00 00 04 D4 B3 04 00 00 00 00 05 D9 B3 02 00 00 00 00 06 89 50 01 00 00 00 00 07 30 5A 01 00 00 00 00 08 30 5A 01 00 00 00 00 09 C4 5D 02 00 00 00 00 0A 90 69 2B 00 00 00 00 0B D1 B3 11 00 00 00 00 0C 04 7D 1B 00 00 00 00 0D B3 65 04 00 00 00 00 0E F5 65 12 00 00 00 00 0F BD B3 1C 00 00 00 00 10 51 C3 04 00 00 00 00 11 F2 65 0A 00 00 00 00 12 AE 65 06 00 00 00 00 13 A0 65 1B 00 00 00 00 14 03 66 25 00 00 00 00 16 E0 65 10 00 00 00 00 17 8E 69 06 00 00 00 00 18 DE 65 03 00 00 00 00 19 7B 69 03 00 00 00 00
'                    offset = 3
'                    Do While offset < UBound(din)
'                        slotno = cInt1(din, offset)
'                        itemid = cInt2(din, offset + 1)
'                        num = cInt1(din, offset + 3)
'                        fine = cInt1(din, offset + 4)
'                        un = a2hex(Mid(hexstring, offset + 8, 8), 8)
'                        'Debug.Print slotno & "," & itemid
'
'                            Me.BackPack.Item(slotno).slot = slotno
'                            Me.BackPack.Item(slotno).itemid = itemid
'                            Me.BackPack.Item(slotno).num = num
'                            Me.BackPack.Item(slotno).fine = fine
'
'                        offset = offset + 12
'                    Loop

'27 02 0A 48 6F 6D 65 2D 41 6C 6F 6E 65 03 4B 8A     'Home-AloneK?
'D6 05 00 07 42 95 77 61 72 95 53 9F 04 01 00 E7     ? B•war•S? ?
'A9 7D 1A 3E AE 7D 1A 00 CB 6E 00 00 01 00 00 00     ?}>?} ?n  
'00 27 7E 09 00 08 FF A7 C1 A7 D2 C2 FB 20 90 01      '~  ??????? 
'01 0A 7D 39 51 07 4C 24 83 20 01 2D 0C 00 00 01     }9QL$? -  
'00 00 00 00 D4 C5 02 00 06 AB D2 B9 BF A7 20 94         ?? ????? ”
'03 01 09 3B E7 07 1F E9 31 C8 0E 01 20 09 00 00       ;??1?
'01 00 00 00 00 4E D7 07 00 05 6F D2 A9 D5 20 7C         N? o??? |
'04 01 07 9C DC 41 21 3E AE 7D 1A 01 89 7B 00 00     ??A!>?}?{
'01 00 00 00 00 92 51 01 00 06 F0 CD D2 A9 D5 F0         ’Q ??????

                'PrintMail(mtype As Byte, uid, msg As String, time As Date)
                Case &H5 'Mail Army
                    ktype = 1
                    kuid = cInt4(din, 3)
                    kmsg = cStrN(din, 15, UBound(din) - 16)
                    kdate = CDate(Hex2Double(Mid(bytes2hexStr(din), 13, 16)))
                    Call PrintMail(ktype, kuid, kmsg, kdate)
                Case &H9
                'Debug.Print "27 09"
                '32 3A 07 00 BA 00 00 00 0D 3A 48 65 4C 4C 42 6C 75 65 42 6F 59 3A
                '88 57 0A 00 BA 00 00 00 0D 3A 48 65 4C 4C 42 6C 75 65 42 6F 59 3A
                Case &H20
                    SendEnd
                Case &H32
                    jobType = cInt1(din, 3)
                    jobPoint = cInt4(din, 4)
                    Select Case jobType
                        Case 1
                            jobName = "เหมืองแร่"
                        Case 2
                            jobName = "ป่าไม้"
                        Case 3
                            jobName = "ปศุสัตว์"
                        Case 4
                            jobName = "กสิกรรม"
                        Case 5
                            jobName = "ก่อสร้าง"
                    End Select
                    Form1.AppendDisplay "มีการพัฒนาอาชีพ " & jobName & " แต้มที่ " & jobPoint, &H10AA10
                Case &H33
                    jobType = cInt1(din, 3)
                    jobPoint = cInt4(din, 4)
                    jobLvl = cInt1(din, 8)
                    Select Case jobType
                        Case 1
                            jobName = "เหมืองแร่"
                        Case 2
                            jobName = "ป่าไม้"
                        Case 3
                            jobName = "ปศุสัตว์"
                        Case 4
                            jobName = "กสิกรรม"
                        Case 5
                            jobName = "ก่อสร้าง"
                    End Select
                    Form1.AppendDisplay "เริ่มการทำอาชีพ " & jobName & " ที่ระดับ " & jobLvl & " แต้มที่ " & jobPoint, &H10AA10
                    senddata MakePacket("F4 44 03 00 20 02 23")
                Case &H36     'ex. 27 36 02 43 01 00 00 00
                    jobPoint = cInt4(din, 4)
                    Form1.AppendDisplay "มีการพัฒนาการทำยุทโธปกรณ์ แต้มที่ " & jobPoint, &H10AA10
                Case &H37     'ex. 27 37 02 41 01 00 00 04 00 00 00
                    jobPoint = cInt4(din, 4)
                    jobLvl = cInt1(din, 8)
                    Form1.AppendDisplay "เริ่มการทำยุทโธปกรณ์ที่ระดับ " & jobLvl & " แต้มที่ " & jobPoint, &H10AA10
                    senddata MakePacket("F4 44 03 00 20 02 23")
            End Select
            
        Case &H32
'F4 44 48 00 32 01
'21 00
'03 02 E3 2E 03 03 00 02 01 00 01 19 5D 00 01
'                  00 01 01 00 01 19 5D 00 01
'                  00 03 01 00 01 19 53 00 01
'21 00
'02 02 E3 2E 03 03 00 02 01 00 01 19 24 01 01
'                  00 01 01 00 01 19 25 01 01
'                  00 03 01 00 01 19 F9 00 01

            offset = 3
                Set oNPCCombayDmg = New Scripting.Dictionary
            Do While offset < UBound(din)
                Set dmgsource = New DamageInfo
                
                fr = cInt1(din, offset + 2)
                fc = cInt1(din, offset + 3)
                fskill = cInt2(din, offset + 4)
                
                dmgsource.AttkFromRow = fr
                dmgsource.AttkFromCol = fc
                dmgsource.AttkSkill = fskill
                
                
                
                n_target = cInt1(din, offset + 6)
                n_count = cInt1(din, offset + 7)
                'MsgBox fr & "," & fc & " - " & fskill & " : " & n_count
                soffset = offset + 8
                For i = 1 To n_count
                    Set DmgTarget = New DamageTarget
                        tr = cInt1(din, soffset)
                        tc = cInt1(din, soffset + 1)
                        statusc = cInt1(din, soffset + 4)
                        DmgTarget.StatusCount = statusc
                        DamagePoint = cInt2(din, soffset + 6)
                        DmgTarget.Status1 = cInt1(din, soffset + 5)
                        DmgTarget.Row = tr
                        DmgTarget.Col = tc
                        DmgTarget.DamagePoint = DamagePoint
                        DmgTarget.Sign1 = cInt1(din, soffset + 8)
                        If DmgTarget.StatusCount = 2 Then
                            DmgTarget.Status2 = cInt1(din, soffset + 9)
                            DmgTarget.DamagePoint2 = cInt1(din, soffset + 10)
                            DmgTarget.Sign2 = cInt1(din, soffset + 12)
                        End If
                        'MsgBox tr & "," & tc & " dmg = " & dmg
                        If godquestion = True Then
                               If tr = Me.Character.Row And tc = Me.Character.Col And DamagePoint = 9999 Then
                                     RaiseEvent onGotGhost
                               ElseIf tr = Me.Character.Row And tc = Me.Character.Col And DamagePoint = 0 Then
                                     RaiseEvent onGotLuckyGod
                               End If
                        End If
                    
                    
                        If fr = Me.Character.Row And fc = Me.Character.Col Then
                            RaiseEvent onSendAttack(fr, fc, tr, tc, fskill)
                        End If
                        If fr = Me.CurrentPartner.Row And fc = Me.CurrentPartner.Col Then
                            RaiseEvent onSendAttack(fr, fc, tr, tc, fskill)
                        End If
                    
                    
                        dmgsource.DmgTarget.Add i, DmgTarget
                    Set DmgTarget = Nothing
                    
                    For n = 0 To Me.oNPCCombat.Count - 1
                        r = Me.oNPCCombat.Item(n).Row
                        c = Me.oNPCCombat.Item(n).Col
                        If r = tr And c = tc Then
                               Me.oNPCCombat.Item(n).HP = Me.oNPCCombat.Item(n).HP - DamagePoint
                        End If
                    Next
                    
                    soffset = soffset + 9
                    soffset = soffset + (statusc - 1) * 4
                    
                Next
                offset = soffset
                oNPCCombayDmg.Add oNPCCombayDmg.Count, dmgsource
                Set dmgsource = Nothing
            Loop
            RaiseEvent onDamage
            RaiseEvent onChangeStatus
        Case &H34
            cmdtype = cInt1(din, 2)
            Select Case cmdtype
                Case &H1    '// Request combat action
                    fightingstate = True
                'RaiseEvent OpenCombatScene
                    RaiseEvent RecvNPCCombat
                    RaiseEvent onOpenCombat

                    If godquestion = False Then
                        If Me.Character.HP > 0 Then
                            RaiseEvent MyAttack
                        End If
                        If Me.CurrentPartner.IsSleep = False And Me.CurrentPartner.HP > 0 Then
                            RaiseEvent PartnerAttack
                        End If
                    Else
                        RaiseEvent ResponseAnswer
                    End If
            End Select
        Case &H35
        'F4 44 04 00 35 03 00 01
        
        Case &HC7
            'F4 44 xx xx C7 06 D1 00 00 00 CA B6 D2 B9 B7 D5 E8 B5 D2 C1     ??   ??????????
            'A4 D3 E3 BA E9 C1 D5 AA D7 E8 CD C7 E8 D2 CD D0     ????????????????
            'E4 C3 3F 09 C1 D5 C1 E9 D2 A2 D2 C7 CD C2 D9 E8     ??? ????????????
            '20 32 20 B5 D1 C7 B9 D0 00   2 ?????
            
            SenderId = cInt1(din, 3)
            MessageText = cStrN(din, 7, UBound(din) - 8)
            Form1.AppendChat "เทพสวรรค์" & SenderId & " : " & MessageText, vbRed
            
            DoEvents
        
        Case Else
 
    End Select
    Exit Sub
ErrorHandler:   ' Error-handling routine.
   
  ' MsgBox Err.Description & " :" & Err.Source & ":" & Err.number & ":"
   Debug.Print Err.Description & " :" & Err.Source & ":" & Err.number & ":"
'   Resume   ' Resume execution at same line
                ' that caused the error.
                RaiseEvent TsError
    Exit Sub
End Sub

Public Sub Disconect()
   ' Debug.Print mvarobj.State
    If mvarobj.state <> 0 Then
        mvarobj.Close
        
        Me.ConnectionState = mvarobj.state
        Set mvarCharacter = Nothing
        Set ol = Nothing
        Set MyPartners = Nothing
        Set mvarNPCSCombat = Nothing
            dBuffer = ""
        
        RaiseEvent Closed
    End If
End Sub
Private Sub mvarobj_Close()
      Debug.Print "Close event raise"
    RaiseEvent Closed
End Sub

Private Sub mvarobj_Connect()
    Me.ConnectionState = mvarobj.state
    RaiseEvent Connected
End Sub
Private Sub mvarobj_DataArrival(ByVal bytesTotal As Long)
    Dim din() As Byte
    If mvarobj.state = sckClosed Then
        Exit Sub
    End If
    mvarobj.GetData din, , bytesTotal
    Call packetBytes2StringAndDecode(din, bytesTotal)
End Sub

Public Sub Reconnect()
    Call Me.Disconect
    Call Form1.DoLogin
End Sub

Private Sub mvarobj_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'mvarobj.Close
    'Me.Disconect
    'Form1.AppendDisplay "SOCKET ERROR #" & Number & " : " & Description, vbRed
    'Form1.AppendDisplay Scode & Source, vbRed
    'SOCKET ERROR #10053 : Connection is aborted due to timeout or other failure
    '-2146818235C:\WINDOWS\system32\MSWINSCK.OCX
    Select Case number
    Case 10053
        Form1.AppendDisplay "Socket error : การเชื่อมต่อถูกยกเลิกหรือใช้เวลาเกินกว่าปกติ !!", vbRed
        Me.Disconect
    Case 10065
        Form1.AppendDisplay "Socket error : ไม่สามารถเชื่อต่อกับ Server ได้ !!", vbRed
        Me.Disconect
    Case Else
        Form1.AppendDisplay "Socket error #" & number & " : " & Description, vbRed
        Me.Disconect
    End Select
    
    RaiseEvent ConnectionError(number, Description, Scode, Source)
End Sub



