/******************************************
Common Function Script (Truebot) 
 ******************************************/
var QA = new ActiveXObject("Scripting.Dictionary"); 
var Wrong = new ActiveXObject("Scripting.Dictionary"); 
//Chat.obj.backColor = 0x0
function ChatClear(){
	Chat.obj.text = "";
}
function DisplayClear(){
	Display.text = "";
}
var PartyFriends = new Array()
var DefaultSena ;
function SetPartyFriend(PlayerName){
	index = PartyFriends.length++;
	PartyFriends[index] = PlayerName;
}
function SetSena(PlayerName){
	DefaultSena = PlayerName;
}
function MonsterAlive(){
	n = 0
	for(i=0;i<ts.oNPCCombat.Count;i++){
		onpc = ts.oNPCCombat.Item(i)
		if(onpc.HP > 0 ){
			n++;
		}
	}
	return n;
}
function findMonster(){
	maxhp = 0
	mi = 0
	for(i=0;i<ts.oNPCCombat.Count;i++){
		onpc = ts.oNPCCombat.Item(i)
		if(onpc.HP > 0 ){
			if(onpc.MAXHP > maxhp){
				maxhp = onpc.MAXHP
				mi = i
			}
		}
	}
	return ts.oNPCCombat.Item(mi)
}
function findMaxLevelMonster(){
	maxlv = 0
	mi = 0
	for(i=0;i<ts.oNPCCombat.Count;i++){
		onpc = ts.oNPCCombat.Item(i)
		if(onpc.HP > 0 ){
			if(onpc.lv > maxhp){
				maxlv = onpc.lv
				mi = i
			}
		}
	}
	return ts.oNPCCombat.Item(mi);
}
function GetNpcObj(row,col){
	for(i=0;i<ts.oNPCCombat.Count;i++){
		onpc = ts.oNPCCombat.Item(i);
		if(onpc.Row == row && onpc.Col == col){
			return onpc;				
		}
	}
	return findMonster();
}
function SelectF1Target(){
	objPos  = new Array()
	objPos[0] = 0;
	objPos[1] = 0;
	maxLevel = 0;
	monsterpattern = new Array()
	monsterpattern[0] = 0;
	monsterpattern[1] = 0;
	for(i=0;i<ts.oNPCCombat.Count ;i++){
		onpc = ts.oNPCCombat.Item(i);
			if(onpc.Row == 0 && onpc.HP > 0){
				monsterpattern[0] += Math.pow(2,onpc.Col); 
			}else if(onpc.Row == 1 && onpc.HP > 0){
				monsterpattern[1] += Math.pow(2,onpc.Col); 
			}
	}
	for(i=0;i<=1;i++){
			for(j=2;j<=5;j++){
			    ptt = 3 << (j-2);
			    if((monsterpattern[i] & ptt ) == ptt){
					objPos[0] = i;
					objPos[1] = j-1;
					return GetNpcObj(objPos[0],objPos[1]);
			    }
			}	
	}
	return findMonster();
}
function SkillID(skname){
	return SKILL.GetId(skname);
}
function SkillSP(skname){
	return SKILL.GetSP(SKILL.GetId(skname));
}
function FindItemInSlot(ItemName){
	for(var i=1;i<=25;i++){
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid);
			if(oItem.getName() == ItemName){
				return oSlot;
			}
	}
	return false;
}
function FindItemInSlotByQty(ItemName,num){
	for(var i=1;i<=25;i++){
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid);
			if(oItem.getName() == ItemName && oSlot.num >=num){
				return oSlot;
			}
	}
	return false;
}
function WantToSale(itemname,num){
	var s = FindItemInSlotByQty(ItemName,num)
	if(s){
		ts.Sale(s.slot,num)
	}
}


function FindItemContribute(ItemName){
	if(s = FindItemInSlot(ItemName)){
		ts.Contribute(0,s.slot);
		debug("[system] • [บริจาคอัตโนมัติ "+ItemName+" "+s.num+" อัน]",0xC08000);
	}
}
function FindItemDrop(ItemName){
	if(s = FindItemInSlot(ItemName)){
		ts.DropItem(s.slot,s.num);
		debug("[system] • [ทิ้งอัตโนมัติ "+ItemName+" "+s.num+" อัน]",0xC08000);
	}
}
function WarpLink( map1 ,warpid1 , map2 ,warpid2){
	if(ts.Character.mapid == map1){
		ts.Warp(warpid1)
		return
	}else if(ts.Character.mapid == map2){
		ts.Warp(warpid2)
		return
	}
}
function Sit(direction){
	ts.SendAction(45+direction)
}


function ReadFile(Fname){
	var ForReading = 1;
	var fso = new ActiveXObject("Scripting.FileSystemObject"); 
	var a = fso.OpenTextFile(Fname, ForReading);
	var contents = a.ReadAll();
	a.Close(); 
	return contents;
}
// WriteLog("c:\\log.txt",ts.Character.Texp);
function WriteLog(Fname,data){
	var ForAppending = 8;
	var fso = new ActiveXObject("Scripting.FileSystemObject"); 
	var a = fso.OpenTextFile(Fname, ForAppending, false);
	a.WriteLine(data); 
	a.Close(); 
}
function ClearLog(Fname){
	var ForWriting = 2;
	var fso = new ActiveXObject("Scripting.FileSystemObject"); 
	var a = fso.OpenTextFile(Fname, ForWriting, false);
	a.WriteLine(""); 
	a.Close(); 
}

function get_random(min,max)
{
	var ranNum= min + Math.round(Math.random()*(max-min));
	return ranNum;
}

var ans
var ans_index = get_random(1,4);
var known = false
/*
function doRecvQuestion(){
	tmp = ts.LastQuestion;
	tmp = tmp.replace("=?","");
	ans = "" + eval(tmp);
	ans_index = ts.LastAnswers.Item(ans);
}

function ResponseAnswer()
{
	 debug("ResponseAnswer",0)
	 ts.Answer(ans_index);
}
*/

/*
function doRecvQuestion(){
	tmp = ts.LastQuestion;
	debug(ts.LastQuestion);
//	tmp = tmp.replace("=?","");
//	ans = "" + eval(tmp);
//	ans_index = ts.LastAnswers.Item(ans);
}
function ResponseAnswer()
{
	 debug("ResponseAnswer",0)
	 ts.Answer(1);
}
*/





var wx = new Array()
var wy = new Array()
var wsec = new Array()
var windex = 0; 


function LoadWaypoint(Fname){
	var ForReading = 1;
	var fso = new ActiveXObject("Scripting.FileSystemObject"); 
	if(!fso.FileExists(Fname)){
		return -1;
	}
	var f = fso.OpenTextFile(Fname, ForReading);
	while (!f.AtEndOfStream){
		s = f.ReadLine();
		debug(s,0)
		w = s.split(" ");

		wx[wx.length] = w[0]
		wy[wy.length] = w[1]
		wsec[wsec.length] = w[2]
	}
	f.Close( );
}

function random_walk(){
	index = get_random(0,wx.length-1)
	ts.Walk(wx[index],wy[index])
	Timer.Interval = wsec[index]
}

function waypoint_walk(){
	index = (windex++) % (wx.length);
	ts.Walk(wx[index],wy[index])
	Timer.Interval = wsec[index]
	debug("walking");
}

function doEatHP(order,difHp){ 
	for(var i = 1;i<= 25 ;i++){ 
		var oSlot = ts.MyItems.Item(i) 
		var oItem = ITEMS.Item(oSlot.itemid) 

		if (oSlot.itemid == 0){ continue; } 
		if(oItem.isHPItem()){ 
			if (oItem.itemvalue > difHp){ continue; } 
			var eatHpAmt = (difHp - (difHp % oItem.itemvalue)) / oItem.itemvalue 
			if (eatHpAmt> 0){ 
				if (eatHpAmt > oSlot.num){ 
					eatHpAmt = oSlot.num; 
				} 
				ts.EatItem(i,eatHpAmt,order);
				debug(oItem.itemname + "  HP  " + oItem.itemvalue + " at slot " + i + "  decrease " + eatHpAmt ,0xC08008 );
				difHp = difHp - eatHpAmt * oItem.itemvalue 
			} 
		} 
	} 
} 

function doEatSP(order,difSp){ 
	for(var i = 1;i<= 25 ;i++){ 
		var oSlot = ts.MyItems.Item(i) 
		var oItem = ITEMS.Item(oSlot.itemid) 
	
		if (oSlot.itemid == 0){ continue; } 
		if(oItem.isSPItem()){ 
			if (oItem.itemvalue > difSp){ continue; } 
			var eatSpAmt = (difSp - (difSp % oItem.itemvalue)) / oItem.itemvalue;
			if (eatSpAmt> 0){ 
				if (eatSpAmt > oSlot.num){ 
					eatSpAmt = oSlot.num; 
				} 
				ts.EatItem(i,eatSpAmt,order) 
				debug(oItem.itemname+"  SP+"+oItem.itemvalue+" at slot"+i+"  decrease"+eatSpAmt ,0xC08008 ) 
				difSp = difSp - eatSpAmt * oItem.itemvalue 
			} 
		} 

		if(oItem.itemtype2 == "SP"){ 
			if (oItem.itemvalue2 > difSp){ continue; } 
			var eatSpAmt = (difSp - (difSp % oItem.itemvalue2)) / oItem.itemvalue2;
			if (eatSpAmt> 0){ 
				if (eatSpAmt > oSlot.num){ 
					eatSpAmt = oSlot.num; 
				} 
				ts.EatItem(i,eatSpAmt,order) 
				debug(oItem.itemname+"  SP+"+oItem.itemvalue2+" at slot"+i+"  decrease"+eatSpAmt ,0xC08008 ) 
				difSp = difSp - eatSpAmt * oItem.itemvalue2
			} 
		} 

	} 
} 

function CheckDisconnect(){ 
	if(ts.Character.HP< (DisconnectFlag * ts.Character.MAXHP)){   // ถ้าเลือดเราน้อยกว่าแค่นี้ให้ตัด
		frm.mnuEnableReconnect.Checked = false; 
		ts.Disconect(); 
		debug("Disconnected : ผู้เล่นเลือดน้อยกว่าที่กำหนด !!",0x0000FF); 
	}

	if(ts.CurrentPartner.HP< (DisconnectFlag * ts.CurrentPartner.MAXHP)){   // ถ้าเลือดขุนน้อยกว่าแค่นี้ให้ตัด
		frm.mnuEnableReconnect.Checked = false; 
		ts.Disconect(); 
		debug("Disconnected : ขุนพลเลือดน้อยกว่าที่กำหนด !!",0x0000FF); 
	}

	if(ts.CurrentPartner.fai < DisconFai){   // ถ้า Fai ต่ำกว่าแค่นี้ให้ตัด
		frm.mnuEnableReconnect.Checked = false; 
		ts.Disconect(); 
		debug("Disconnected : ขุนพลซื้อสัตย์ต่ำกว่าที่กำหนด !!",0x0000FF); 
	}
} 

function onAnswerWrong(q,a){ 
} 

function FinishAnswerFuckGod(){ 
	if(ghost_count>=3){
		ts.Disconect();
	}
}

function BasicAttack(Char, Target, Skill){ 
	if (Char.SP >= SkillSP(Skill)){
		ts.SendAttack(Char.Row, Char.Col, Target.Row, Target.Col, SkillID(Skill)) 
	} else {
		ts.SendAttack(Char.Row, Char.Col, Target.Row, Target.Col, SkillID("มือเปล่า")) 
	}
} 

function BattleStarted(){ 
	battle_count++
	roundcount = 0;
} 

function PartyStop( playerid ){ 
} 



debug("1. Common.js  -- loaded successful !!" , 0x00AA00);