Attribute VB_Name = "ModMain"
Public Sub Main()
   ' MsgBox fnTrueDLL()
   ' Form2.Show 1
      'InitCommonControlsVB
    Set dicExp1 = New Scripting.Dictionary
    SetExpNormal
    Set dicExp2 = New Scripting.Dictionary
    SetExpNewBorn
    Set dicExp3 = New Scripting.Dictionary
    SetExpNewBorn2
    Set ditems = New Scripting.Dictionary
    'Call LoadItemData
    
    Set dskills = New Scripting.Dictionary
    'Call LoadSkillData
    
    Set dmaps = New Scripting.Dictionary
    LoadMaps
    
    Set dnpcs = New Scripting.Dictionary
    'Call LoadNPCData

    LastSelectItem = 1
    LastSelectBPItem = 1
    Form1.Show
   ' MDIForm1.Show
End Sub
