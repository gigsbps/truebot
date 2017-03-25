function SetContributeItemList(){ 
	AddContributeItemList("เม็ดหินกบ") 
	AddContributeItemList("เม็ดหินสาวงาม") 
}

//************** ตั้งรายชื่อ ทิ้งของ*****************
function SetDropItemList(){ 
	AddDropItemList("ต้านปล่อยพิษ") 

} 

function AutoSendItems(){ 
	SendingAllSet(896000,"หินถามฟ้า");
	SendingAllSet(896000,"หินเขียวคู่");
	SendingAllSet(896000,"หินกบ");
}

debug("4. Drop_Con.js  -- loaded successful !!" , 0x00AA00);