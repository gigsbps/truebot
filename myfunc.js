/******************************************
My Private Function Script (Truebot) 
 ******************************************/

function FixTypeItemBug(){
	for(var i=1;i<=25;i++){
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid);
			if (oItem.getName() == "ยาคืนวิญญาณ"){
				oItem.itemtype = "";
			}
			if (oItem.itemvalue2 < 0){
				oItem.itemvalue2 = 0;
			}
			//debug("Slot " + i + " : " +oItem.getName() + " " + oItem.itemtype + "=" + oItem.itemvalue + " " + oItem.itemtype2 + "=" + oItem.itemvalue2)
	}
}

function SendingDupeSet(SendingName,SendingItem){
	var Sending = 0;
	var MyInventory = new Array();
	for(var i=1;i<=25;i++){ 
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid); 
		var itemname = oItem.getName();
		MyInventory[i] = itemname;
		if ((Sending == 0) && (itemname == SendingItem)){
			debug("Found item : " + SendingItem + " = " + ts.MyItems(i).num + " in slot " + i);
			for(var k=1;k < i;k++){ 
				if ((MyInventory[k] == MyInventory[i]) && ((ts.MyItems(k).num==50) && (ts.MyItems(i).num==50))){
					Sending = 1;
					ts.SendItemTo(SendingName,oSlot.slot,oSlot.num);
					debug("Sending item : " + itemname + " to " + SendingName);
				}
			}
		}
	} 
}

function SendingAllSet(SendingName,SendingItem){
	var Sending = 0;
	var MyInventory = new Array();
	for(var i=1;i<=25;i++){ 
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid); 
		var itemname = oItem.getName();
		if ((Sending == 0) && (itemname == SendingItem) && (ts.MyItems(i).num==50)){
			Sending = 1;
			ts.SendItemTo(SendingName,oSlot.slot,oSlot.num);
			debug("Sending item : " + itemname + " to " + SendingName);
		}
	} 
}

function SendingByName(SendingName,SendingItem){
	var Sending = 0;
	var MyInventory = new Array();
	for(var i=1;i<=25;i++){ 
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid); 
		var itemname = oItem.getName();
		if ((Sending == 0) && (itemname == SendingItem)){
			Sending = 1;
			ts.SendItemTo(SendingName,oSlot.slot,oSlot.num);
			debug("Sending item : " + itemname + " to " + SendingName);
		}
	} 
}

function ContributingDupeSet(ContributingItem){
	var MyInventory = new Array();
	for(var i=1;i<=25;i++){ 
		var oSlot = ts.MyItems(i); 
		if( oSlot.itemid == 0){ continue; } 
		var oItem = ITEMS.Item(oSlot.itemid); 
		var itemname = oItem.getName();
		MyInventory[i] = itemname;
		if (itemname == ContributingItem){
			debug("[system] • [Found item : " + ContributingItem + " = " + ts.MyItems(i).num + " in slot " + i +"]",0xC08000);
			for(var k=1;k < i;k++){ 
				if (ts.MyItems(k).num >=30){
					ts.Contribute(0,oSlot.slot);
					debug("[system] • [บริจาคอัตโนมัติ "+itemname+" "+ts.MyItems(i).num+" อัน]",0xC08000);
				}
			}
		}
	} 
}

//*********** ฟั่งชั่นเกี่ยวกับการ ทิ้ง+บริจาคของ**********

var DropItemList = new Array(); 
var ContributeItemList = new Array(); 
var dItemIndex = 0; 
var cItemIndex = 0; 

function AddContributeItemList( ItemName ){ 
	ContributeItemList[cItemIndex++] = ItemName 
} 

function AutoContributeItems(){ 
	for(var i=0;i<ContributeItemList.length;i++){
		FindItemContribute(ContributeItemList[i]);
	}
} 

function AddDropItemList( ItemName ){ 
	DropItemList[dItemIndex++] = ItemName 
} 

function AutoDropItems(){ 
	for(var i=0;i<DropItemList.length;i++){
		FindItemDrop(DropItemList[i]);
	}
} 

//*******************************************************

//**************** ฟั่งชั่นเกี่ยวกับการกินของ*****************
function Heal(){ 
var skillHeal = "";
var nameHeal = "";
var SpForHeal = 0;
	switch (skillHealId) {
		case 1:  skillHeal = 11004;  SpForHeal = 22;  nameHeal = "วารีคืนพลังนอกฉาก"; break;
		case 2:  skillHeal = 11007;  SpForHeal = 35;  nameHeal = "รักษาบาดเจ็บนอกฉาก"; break;
		case 3:  skillHeal = 11010;  SpForHeal = 42;  nameHeal = "สุดยอดการรักษานอกฉาก"; break;
	}

	if((skillHealId >= 1) && (skillHealId <=3) && (ts.Character.SP >= SpForHeal)){
		while((ts.Character.HP < (ts.Character.MAXHP - HealAmt)) && (ts.Character.SP >= SpForHeal)){ 
	  		ts.Heal(skillHeal,ts.Character.uid);    
			debug(nameHeal + "ตัวเรา" , 0) 
			frm.cdelay(1); // เหมือน ts.delay(1000) แต่ไม่ lag
	  	} 
	 	while((ts.CurrentPartner.HP < (ts.CurrentPartner.MAXHP - HealAmt)) && (ts.Character.SP >= SpForHeal)){ 
 			ts.HealPartner(skillHeal,1); 
			debug(nameHeal + "ขุนพล" , 0) 
			frm.cdelay(1);
	  	} 
		
	}	
}

function AutoEatFood(){
	Heal();
	if (ts.Character.HP < (ts.Character.MAXHP * hpFractionEat)){ 
		debug("ผู้เล่นกิน HP" , 0);
       		doEatHP(0,(ts.Character.MAXHP * hpFraction)-ts.Character.HP);
  	} 
    	if (ts.Character.SP < (ts.Character.MAXSP * spFractionEat)){ 
		debug("ผู้เล่นกิน SP" , 0);
         		doEatSP(0,(ts.Character.MAXSP * spFraction)-ts.Character.SP);
   	} 
     	if (ts.CurrentPartner.HP < (ts.CurrentPartner.MAXHP * hpFractionEat)){ 
		debug("ขุนพลกิน HP" , 0);
       		doEatHP(ts.CurrentPartner.Order,(ts.CurrentPartner.MAXHP * hpFraction)-ts.CurrentPartner.HP);
  	} 
    	if (ts.CurrentPartner.SP < (ts.CurrentPartner.MAXSP * spFractionEat)){ 
		debug("ขุนพลกินSP" , 0);
       		doEatSP(ts.CurrentPartner.Order,(ts.CurrentPartner.MAXSP * spFraction)-ts.CurrentPartner.SP);
  	} 
}
//*******************************************************

function Avoid9am(){
	var time = new Date();
	h = time.getHours();
	m = time.getMinutes();
	
	if((h == 8) && (m >= 50)){
		debug("ขณะนี้เวลา 8:50 นาฬิกา ตัดสาย 20 นาทีเพื่อนหลบช่วงบำรุง server");
		ts.Disconect(); // ตัดสายตัวเอง
		frm.cdelay(20*60);//หน่วงเวลาการล๊อคอิน
		frm.cmdLogin.value=true; 
	} else {
		debug("ขณะนี้เวลา " + h + ":" + m + " นาฬิกา");
	}
}

function Move(MapX,MapY)
{
	if ((ts.Character.x!=MapX) && (ts.Character.y!=MapY))
	{
		debug("[system] • [Walk ("+ts.Character.x+","+ts.Character.y+")>>("+MapX+","+MapY+")]",0xC08000);
		ts.Walk(MapX,MapY);
	}
}

function SwapLucky(taketype){ 
	var slotno = 25; 
	if (taketype == "Takeon") { 
		ts.Equipment(slotno); 
		LuckyStatus = 1; 
		debug("************************************",0xFF0000) 
		debug("ใส่ " +ITEMS.Item(ts.MyItems(slotno).itemid).getName() +" เรียบร้อย",0xFF0000) 
		debug("************************************",0xFF0000) 
	} 

	if (taketype == "Takeoff") { 
		if (LuckyStatus == 1) { 
			ts.Equipment(slotno); 
			LuckyStatus = 0; 
			debug("************************************",0xFF0000) 
			debug("ถอด " +ITEMS.Item(ts.MyItems(slotno).itemid).getName() +" เรียบร้อย",0xFF0000) 
			debug("************************************",0xFF0000)			 
		} 

	} 

	if (taketype == "Logon") { 
		if(ts.MyItems(slotno).itemid != 23085 && ts.MyItems(slotno).itemid != 23023){ 
			ts.Equipment(slotno); 
			LuckyStatus = 0; 
			debug("************************************",0xFF0000) 
			debug("ถอด " +ITEMS.Item(ts.MyItems(slotno).itemid).getName() +" เรียบร้อย",0xFF0000) 
			debug("************************************",0xFF0000) 
		}else{ 
			debug("************************************",0xFF0000) 
			debug("ไม่เปลี่ยนแปลง",0xFF0000) 
			debug("************************************",0xFF0000) 
		} 
	} 
	
}

function EatGod(){ 
	for(var i = 1;i<=25 ;i++){ 
		if (CurrentGodNum <= 250) { 
			var oSlot = ts.MyItems.Item(i) 
			if(oSlot.itemid == 0){ continue; } 
			var oItem = ITEMS.Item(oSlot.itemid) 

			if(oItem.getName() == "เทพโชคลาภใหญ่") { 
				NumToEat = Math.floor((250 - CurrentGodNum) / 10);
				if (NumToEat > 0){
					if (oSlot.num > NumToEat)  {
						CurrentGodNum = CurrentGodNum + (10*NumToEat);
						ts.EatItem(i, NumToEat, 0) 
						debug("** กิน"+oItem.getName() + " จำนวน " + NumToEat +" **");
					} else {
						CurrentGodNum = CurrentGodNum + (10*oSlot.num);
						ts.EatItem(i, oSlot.num, 0) 
						debug("** กิน"+oItem.getName() + " จำนวน " + oSlot.num +" **");
					}
				}
			} 

			if (oItem.getName() == "เทพโชคลาภ") { 
				NumToEat = 250 - CurrentGodNum;
				if (NumToEat > 0){
					if ((oSlot.num > NumToEat) && (NumToEat > 0)) {
						CurrentGodNum = CurrentGodNum + NumToEat;
						ts.EatItem(i, NumToEat, 0) 
						debug("** กิน"+oItem.getName() + " จำนวน " + NumToEat +" **");
					} else {
						CurrentGodNum = CurrentGodNum + oSlot.num;
						ts.EatItem(i, oSlot.num, 0) 
						debug("** กิน"+oItem.getName() + " จำนวน " + oSlot.num +" **");
					}
				}
			}
		} 
	}
}

function ViewState(){ 
	debug("************************************",0xFF0000) 
	debug(" Battle Count  : " + battle_count ,0xFF0000) 
	debug(" Warrior's FAI : " + ts.CurrentPartner.CharName +" : " + ts.CurrentPartner.fai,0xFF0000) 
	//debug(" Lucky God Count : " + CurrentGodNum,0xFF0000) 
	//debug(" Slot of current using LuckyBadge : " + LuckySlot,0xFF0000) 
	debug("************************************",0xFF0000) 
} 

function SwapLuckyBadge(){
	LuckySlot = LuckySlot + 1;
	if (LuckySlot > 5){
		LuckySlot = 1;
	}

	ts.Equipment(LuckySlot);
	debug("ใส่ " +ITEMS.Item(ts.MyItems(LuckySlot).itemid).getName() +" เรียบร้อย",0xFF0000) 
}

debug("2. MyFunc.js  -- loaded successful !!" , 0x00AA00);