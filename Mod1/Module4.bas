Attribute VB_Name = "Module4x"
Function hashNumberItem(value)
'if(Answer<0) Answer = Answer + 0xFF
    hashNumberItem = (value Xor &HB4) - 109
End Function


Function ITemTypes(itemtype)
        Select Case itemtype
            Case 24
                Itemty = "Atk"
            Case 202
                Itemty = ""
            Case 26
                Itemty = "MaxHP"
            Case 27
                Itemty = "MaxSP"
            Case 28
                Itemty = "Agi"
            Case 30
                Itemty = "Int"
            Case 31
                Itemty = "Def"
            Case 138
                Itemty = "Loy"
            Case 225
                Itemty = "HP"
            Case 224
                Itemty = "SP"
            Case Else
                Itemty = itemtype
            
        End Select
ITemTypes = Itemty
End Function

Function ITemGroup(itemtype)
        Select Case itemtype
            Case 144
                Itemty = "อาวุธ"
            Case 180
                Itemty = "คัมถีร์"
            Case Else
                Itemty = itemtype
            
        End Select
ITemGroup = Itemty
End Function

Function ITemEQ(itemtype)
        Select Case itemtype
            Case 1
                Itemty = "หมวก"
            Case 2
                Itemty = "เสื้อ"
            Case 3
                Itemty = "อาวุธ"
            Case 4
                Itemty = "แขน"
            Case 5
                Itemty = "เท้า"
            Case 6
                Itemty = "เครื่องประดับ"
            Case Else
                Itemty = itemtype
            
        End Select
ITemEQ = Itemty
End Function


Sub LoadItemData()
Dim bArray() As Byte
Dim offset As Long
Dim datasize As Long
Dim filesize As Long
Dim filename As String
Dim foundDat As Boolean
    
    offset = &H173 + &H172
    datasize = &H172
    foundDat = False
    If Form1.GameFolder <> "" Then
        If Dir(Form1.GameFolder & "\data\Item.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) <> "" Then
            filename = Form1.GameFolder & "\data\Item.Dat"
            foundDat = True
        End If
    End If
    If Not foundDat Then
        If Dir(App.Path & "\Item.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
            If Dir("C:\Program Files\Asiasoft\TSOnline\data\Item.Dat", vbNormal + vbHidden + vbSystem + vbReadOnly) = "" Then
                MsgBox "ไม่สามารถหาไฟล์ Skill.dat ได้", vbExclamation + vbOKOnly, "Error !!"
                Unload Form1
                End
            Else
                filename = "C:\Program Files\Asiasoft\TSOnline\data\Item.Dat"
            End If
        Else
            filename = App.Path & "\Item.Dat"
        End If
    End If
    bArray = ReadFile(filename, filesize)
    
    Index = 0
    Do While offset < filesize
        itemx = ((&HEFC0 Xor cInt2(bArray, offset + &H15) Xor &HFFFF0000) Xor &H3) - 9
        
        itemlimit = cInt1(bArray, offset + &H16)
        
        itemtype = cInt1(bArray, offset + 31)
        itemtype2 = cInt1(bArray, offset + 33)
        unknow1 = cInt1(bArray, offset + 45)
        
        
        itemv = cInt1(bArray, offset + 37)
        itemvalue = hashNumberItem(itemv)
        
        itemv2 = cInt1(bArray, offset + 41)
        itemvalue2 = hashNumberItem(itemv2)
        
        
        itemcontribute = cInt1(bArray, offset + &H35)
        itemcontribute = hashNumberItem(itemcontribute)
        
        itenname = ""
        n = offset + &H13
        Do While n > offset
            If bArray(n) = 0 Then Exit Do
            c = Chr(bArray(n))
'
            itenname = itenname & c
            n = n - 1
        Loop
        Itemty = ITemTypes(itemtype)
        Itemty2 = ITemTypes(itemtype2)
        itendesc = ""
        n = offset + &H170
        Do While True
            If bArray(n) = 0 Then Exit Do
            c = Chr(bArray(n))
'
            itendesc = itendesc & c
            n = n - 1
        Loop
        ITemEQment = cInt1(bArray, offset + 47)
        ITemEQment = (ITemEQment Xor &H9A) - 9
        
        If itemx = 47198 Then itenname = "สร้อยหลี่ซู่ (Int)"
        If itemx = 47229 Then itenname = "สร้อยหลี่ซู่ (Atk)"
        If itemx = 51036 Then itenname = "ดาวหลี่ซู่ (Int)"
        If itemx = 51103 Then itenname = "ดาวหลี่ซู่ (Atk)"
        
        Call SetItems(itemx, itenname, Itemty, itemvalue, Itemty2, itemvalue2, itendesc, ITemEQ(ITemEQment))
        offset = offset + datasize
    Loop
    Call SetItems(0, "", "", 0, "", 0, "", "")
    Call SetItems(99999, "UNKNOW", "", 0, "", 0, "", "")
End Sub
