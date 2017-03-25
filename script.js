// *********************************
//TrueBot script for v3.0.1
//Modify by X CroSs
//Date 24/10/05 : 6.52am
//**********************************
var ghost_count = 0; 
var battle_count = 0; 
var roundcount = 0; 
var party_count = 0;
var state = "";
var DisconnectFlag = 0.3; 

var hpFractionEat = 0.9;	//บอทจะเริ่มกินเมื่อเลือดน้อยกว่า 70%
var spFractionEat = 0.9;
var hpFraction = 0.99;		//บอทจะกินถึง 95% ของเลือดจริง
var spFraction = 0.99;
var skillHealId = 0;		//0,1,2,3  	// 0 ปิดการฮิลนอกฉาก ไม่มีสกิลฮิลไดๆ; 1 วารีคืนพลัง; 2.รักษาบาดเจ็บ;3.สุดยอดรักษา
var HealAmt = 115;		//ความแรงของสกิลฮีล
var DisconFai = 20;		//ถ้าซื่อน้อยกว่านี้ให้ตัด
var avoild9amFlag = 1;	// ตั้งว่าให้ ต่อบอทใหม่ตอน 9 โมงหรือไม่ 1=ต่อ 0=ไม่ต่อ

function MyAttack(){ 
	state="fight"
	var MyChar = ts.Character 
	var Warrior = ts.CurrentPartner 
	//Monster = SelectF1Target() 
	Monster = findMonster()
	var n = MonsterAlive()  
	// if(NPC.Item(onpc.uID).charname != "แร่"){  // ไม่ใช่แร่ให้หนี
	//	BasicAttack(MyChar, Monster, "หลบหนี"); 
	//} else {
		BasicAttack(MyChar, Monster, "มือเปล่า"); 
	//}

	//ts.SendAttack(MyChar.Row, MyChar.Col, 0, 5, SkillID("ฟันใต้ดิน"))
} 

function MyPartnerAttack(){ 
	var MyChar = ts.Character 
	var Warrior = ts.CurrentPartner 
	//Monster = SelectF1Target() 
	Monster = findMonster() 
	var n = MonsterAlive()  
	// if(NPC.Item(onpc.uID).charname != "แร่"){  // ไม่ใช่แร่ให้หนี
	//	BasicAttack(Warrior, Monster, "หลบหนี"); 
	//} else {
		BasicAttack(Warrior, Monster, "มือเปล่า"); 
	//}

}

function BattleStoped(){ 

	//**********ส่งของแบบเกิน 2 ชุด************
	//SendingDupeSet("อัคคีทลายฟ้า1","ลิ้นไซซี");
	// ********* สำคัญมาก ๆ กันขุนหนี ***********
	CheckDisconnect();

	// ******** กิน hp/sp + ฮีลนอกฉาก**********
	AutoEatFood()

	//********** บริจาค + ทิ้งของ **************** 
	AutoDropItems();
	AutoContributeItems();
	AutoSendItems();
	ViewState();
	state = "walk"

	//ts.ClickOnNPC(1)
} 



LoadWaypoint("w1.txt")

function OnTimer() {
	waypoint_walk()
	//debug("x x",0x0000FF)
}

function InitBot(){ 
	// ******** กิน hp/sp + ฮีลนอกฉาก**********
	AutoEatFood();

	//***** ตรวจสอบเลือด + ซื่อลูกก่อนเริ่ม
	CheckDisconnect();

	//***** ตั้งค่ารายชื่อ ของ ดรอป + บริจาค
	SetContributeItemList();
	SetDropItemList();

	// ***** หาก script ถูก จะขึ้น ข้อความนี้เป็นสีเขียว
	debug("5. Script.js  -- loaded successful !!" , 0x00AA00);
	debug("=== Successfull loading all scripts. ===" , 0x0000AA);

} 

InitBot()


