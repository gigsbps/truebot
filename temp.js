function Start()
{
	ts.ClickOnNPC(13);
}

function NpcDialogMenu(DialogId){
	debug("DialogMenu = "+DialogId,0x0000FF);
	ts.SelectChoice(1);
	ts.SendEnd();
}

function InitBot(){ 
	debug("ARMY BOT loaded ^^" , 0x00AA00);
} 

InitBot()
