
function replace_all(str , diml ,rep ){
	try{
         var re = new RegExp(diml,"i");
         while(re.test(str)){
            str = str.replace(re,rep);
         }
	}catch(e){
	}
	return str
}
function findAnswerArray(ansArray){
   a = (new VBArray(ts.LastAnswers.Keys())).toArray();  
    for(var i=0;i<4;i++){
      for(var j=0;j<ansArray.length;j++){
         if(a[i] == ansArray[j]){
            return i+1;
         }
      }
   }   
   return -1;
}
function fineExistingQuestion(){
var reindex= -1;
	if(QA.Exists(ts.LastQuestion)){
		ans = QA.Item(ts.LastQuestion)
		if(typeof(ans) == 'object'){
		   reindex = findAnswerArray( ans )
		}else{
 		   ans2 = ans.replace(" ","-");
		   if(ts.LastAnswers.Exists(ans)){
			  reindex = ts.LastAnswers.Item(ans);

		   }else if(ts.LastAnswers.Exists(ans2)){
			  reindex = ts.LastAnswers.Item(ans2);
		   }else{
		      reindex = -1;
		   }
		}  
    }else{
      reindex = -1;
	}
	return reindex;
}
function how_many(){
	var reindex = -1;
	var sq = ts.LastQuestion
		sq = sq.replace("How many peopleare","How many people are")
	if(sq.indexOf("How many people are there in the group of")!=-1){
       dm = sq.split("How many people are there in the group of ");
       dm[1] = replace_all(dm[1] , "and" ,",");
       dm[1] = replace_all(dm[1] , "," ," and ");
 	   if(dm[1].indexOf(" and ")!=-1){
	      c = dm[1].split(" and ");
		  a = (new VBArray(ts.LastAnswers.Keys())).toArray();  
		  for(var i=0;i<4;i++){
		     if(a[i].indexOf(c.length)!=-1){
				reindex = i+1;
				return reindex;
			 }
		  }
	   }
   }	   
   return -1;
}
function between(){
	var reindex= -1;
	var sq = ts.LastQuestion
	if(sq.indexOf(" between ")!=-1){
            dm = sq.split(" between ");
            d1 = dm[1];
            a1 = d1.split(" and ");
            a = (new VBArray(ts.LastAnswers.Keys())).toArray();  
            for(var i=0;i<4;i++){
               choice = replace_all(a[i]  , " " ,"-");
               ////////////////////////////////////////////////////////////////////
               ans = replace_all(a1[0]    , "  " ," ");
               ans = replace_all(ans      , " " ,"-");
               if(ans.indexOf(choice)!=-1){
                // WriteLog("know.txt","NOT Existing Q (between) Solved: "+choice)
                 return i+1;
               }
               ////////////////////////////////////////////////////////////////////
               ans = replace_all(a1[1]    , "  " ," ");
               if(ans.indexOf(" in ")!=-1){
                  a2 = ans.split(" in ");
                  ans  = a2[0];
               }
               ans = replace_all(ans      , " " ,"-");
               ans = replace_all(ans      , "?" ,"");
               ans = replace_all(ans      , ":" ,"");
               if(ans.indexOf(choice)!=-1){
                 //WriteLog("know.txt","NOT Existing Q (between) Solved: "+choice)
                 return i+1;
               }
               ////////////////////////////////////////////////////////////////////
			}
     }else{
		return -1;
	 }
}
function solve_brother_of(){
var reindex= -1;
	re = /(.*) is the (.*) brother of:/i;
    if(re.test(ts.LastQuestion)){
       r = ts.LastQuestion.match(re);
       a = (new VBArray(ts.LastAnswers.Keys())).toArray();
       for(i = 0 ;i<4;i++){
	       if(r[1].indexOf(" ")!=-1){
			  p = r[1].split(" ")
			  prefix = p[0]
		   }else{
			  prefix = r[1];
		   }
           if(a[i].indexOf(prefix)!=-1){
              return reindex = i+1;
           }
      }
   }
   return reindex;
}
function solve_father(){
var reindex= -1;
   re = /(.*) is the father (.*)/i;
   if(re.test(ts.LastQuestion)){
      r = ts.LastQuestion.match(re);
      a = (new VBArray(ts.LastAnswers.Keys())).toArray();
      for(i = 0 ;i<4;i++){
		  if(r[1].indexOf(" ")!=-1){   
             p = r[1].split(" ")
             prefix = p[0];
          }else{
             prefix = r[1]; 
          }
          if(prefix.indexOf(a[i])!=-1){
            return reindex = i+1;
          }
      }
   }
   return reindex;
}
function solve_father2(){
var reindex= -1;
   re = /Who is father of (.*)/i;
   if(re.test(ts.LastQuestion)){
      r = ts.LastQuestion.match(re);
      a = (new VBArray(ts.LastAnswers.Keys())).toArray();
      for(i = 0 ;i<4;i++){
		  if(r[1].indexOf(" ")!=-1){   
             p = r[1].split(" ")
             prefix = p[0];
          }else{
             prefix = r[1]; 
          }
          if(a[i].indexOf(prefix)!=-1){
            return reindex = i+1;
          }
      }
   }
   return reindex;
}
function guess_strategist(){
// "Among the followings, who was the strategist of Cao Cao?:"
	var reindex= -1;
	re = /Among the followings, who was the strategist of (.*)?:/i;
    if(re.test(ts.LastQuestion)){
       r = ts.LastQuestion.match(re);
       a = (new VBArray(ts.LastAnswers.Keys())).toArray();
       for(i = 0 ;i<4;i++){
	       if(a[i].indexOf("-")!=-1){
              return i+1;
           }
       }
   }
   return reindex;
}
function guess_not_strategist(){
//Among the followings, who was not the strategist of Cao Cao?:
	var reindex= -1;
	re = /Among the followings, who was not the strategist of (.*)?:/i;
    if(re.test(ts.LastQuestion)){
       r = ts.LastQuestion.match(re);
       a = (new VBArray(ts.LastAnswers.Keys())).toArray();
       for(i = 0 ;i<4;i++){
	       if(a[i].indexOf("-")==-1){
              return i+1;
           }
       }
   }
   return reindex;
}
function general_come_from(){ 
	var reindex= -1;
	re = /Where does(.*)general come from(.*)/i;

    if(re.test(ts.LastQuestion)){
       r = ts.LastQuestion.match(re);
       a = (new VBArray(ts.LastAnswers.Keys())).toArray();
       for(i = 0 ;i<4;i++){
      if(a[i].indexOf("-")!=-1){
              return i+1;
           }
       }
   }
   return reindex;
}
function position(){
var reindex= -1;
   re = /(.*)what is the position(.*)/i;
   if(re.test(ts.LastQuestion)){
      r = ts.LastQuestion.match(re);
      a = (new VBArray(ts.LastAnswers.Keys())).toArray();
      for(i = 0 ;i<4;i++){
          if(a[i].indexOf("General")!=-1){
            return reindex = i+1;
          }
      }
   }
   return reindex;
}
function surname(){
var reindex= -1;
   re = /(.*)had the surname of:/i;
   if(re.test(ts.LastQuestion)){
      r = ts.LastQuestion.match(re);
      a = (new VBArray(ts.LastAnswers.Keys())).toArray();
      for(i = 0 ;i<4;i++){
          if(r[1].indexOf(a[i])!=-1){
            return reindex = i+1;
          }
      }
   }
   return reindex;
}
var ans
var ans_index = -1;
function doRecvQuestion(){
	try{
		if((ans_index=fineExistingQuestion())!=-1) { 
		    WriteLog("know.txt","Solved by QA Existing. :"+ts.LastQuestion)
			return 
		}
		if((ans_index=how_many())!=-1) { 
		    WriteLog("know.txt","Solved by (How many) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=between())!=-1) { 
		    WriteLog("know.txt","Solved by (between) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=solve_brother_of())!=-1) { 
		    WriteLog("know.txt","Solved by (brother_of) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=solve_father())!=-1) { 
		    WriteLog("know.txt","Solved by (father_of) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=solve_father2())!=-1) { 
		    WriteLog("know.txt","Solved by (solve_father2) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=guess_strategist())!=-1) { 
		    WriteLog("know.txt","Solved by (guess_strategist) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=guess_not_strategist())!=-1) { 
		    WriteLog("know.txt","Solved by (guess_not_strategist) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=general_come_from())!=-1) { 
		    WriteLog("know.txt","Solved by (general_come_from) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=position())!=-1) { 
		    WriteLog("know.txt","Solved by (position) :"+ts.LastQuestion)
			return 
		}
		if((ans_index=surname())!=-1) { 
		    WriteLog("know.txt","Solved by (surname) :"+ts.LastQuestion)
			return 
		}
			 
		
		ans_index = get_random(1,4)
	    WriteLog("know.txt","Last way (random it.) :"+ts.LastQuestion)
	}catch(e){
	}
}



function ResponseAnswer()
{
	 debug("PH-ResponseAnswer",0)
 	 a = (new VBArray(ts.LastAnswers.Keys())).toArray();
     WriteLog("know.txt","--------------------------------------------------------------")
	 WriteLog("know.txt",ts.LastQuestion)
     for(i = 0 ;i<4;i++){
		WriteLog("know.txt","Choice is (\""+ a[i] +"\")")
	 }
	 debug("ans_index = "+ans_index,0x00FF00);
	 ts.Answer(ans_index);
	 debug("Current Answer = "+ts.LastResponseAnswer,0x00FF00);
	 WriteLog("know.txt","AddQA(\""+ts.LastQuestion+"\",\""+ts.LastResponseAnswer+"\");")
     WriteLog("know.txt","==============================================================")
}




function onAnswerRight(Question,Ans){
	if(!QA.Exists(Question)){
		WriteLog("../QA.js","AddQA(\""+Question+"\",\""+Ans+"\");")
		debug("Record new QA.",0x0000FF)
	}else{
		debug("QA. Existing",0x0000FF)
		if(Ans!=QA.Item(Question)){
			WriteLog("../QA.js","//Remark Duplicate : AddQA(\""+Question+"\",\""+Ans+"\");")
			debug("QA. Ans Duplicate.",0x0000FF)
		}
	}
}
function AddQA(q,a){
	try{
		QA.Add(q,a)
	}catch(e){
		debug("error : "+(QA.Count) ,0x000FF)
	}
}
