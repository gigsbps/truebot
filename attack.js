function BattleStarted()
{ 	
	PartyCheck("Start");
	PartyCheck("Stop")
	battle_count++
	roundcount = 0;
	Turn++
		if (Turn == 10)
		{Turn = 0	}
} 


function MyAttack()
{	Fighting = 1
	LagTime = 0
	PartyCheck("Stop");
		Battle++
    	m = findMonster(); 
    	n = MonsterAlive(); 
			
				if (n == 6){ sk = SkillID("ธนูไฟ");
 ts.SendAttack( ts.Character.Row , ts.Character.Col , 0 , 2, sk  ); }
	else						{sk = SkillID("มือเปล่า");  
ts.SendAttack( ts.Character.Row , ts.Character.Col , 1 , 2, sk  ); }
    ts.SendAttack( ts.Character.Row , ts.Character.Col , m.Row , m.Col , sk  ); 

} 


function MyPartnerAttack(){ 
   	m = findMonster(); 
   	n = MonsterAlive(); 
		
				if (n == 6){ sk = SkillID("ธนูไฟ");
				  ts.SendAttack( ts.CurrentPartner.Row , ts.CurrentPartner.Col , 0 , 2 , sk ); }
	else						{sk = SkillID("มือเปล่า"); 
	 ts.SendAttack( ts.CurrentPartner.Row , ts.CurrentPartner.Col , 1 , 2 , sk ); }

   ts.SendAttack( ts.CurrentPartner.Row , ts.CurrentPartner.Col , m.Row , m.Col , sk ); 

} 


function BattleStoped()
{	Fighting = 0
	Battle = 0
	AutoEatFood()
	AutoContributeItems()

	AutoDropItems()
	ItemSend()
	/*************** สำคัญมาก ๆ กันขุนหนี ************/
	CheckDisconnect();
	ViewState();
	PartyCheck("Stop");
	Battlecount()
	if (Turn == 3){SwapLuckyBadge()	}

} 


	debug("ไฟล์ที่ 5. attack.js  -- รันสำเร็จ !!" , 0xFF5588);