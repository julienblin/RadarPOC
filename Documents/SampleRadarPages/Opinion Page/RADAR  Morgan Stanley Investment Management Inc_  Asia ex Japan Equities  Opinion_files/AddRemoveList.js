
function handleEnterPress(callIfEnter){	
	var Char = event.keyCode;
	if (Char == 13){
		eval(callIfEnter);
		window.event.cancelBubble=true;
		window.event.returnValue=false;		
	}
}

function setEnableDisable(cId,disabled){
	leftList = document.getElementById(cId + '_leftList');
	leftList.disabled = disabled;
	
	rightList = document.getElementById(cId + '_rightList');
	rightList.disabled = disabled;
	
	oneToLeftButton = document.getElementById(cId + '_otlImg');
	oneToLeftButton.disabled = disabled;
	
	oneToRightButton = document.getElementById(cId + '_otrImg');
	oneToRightButton.disabled = disabled;
	
	allToRightButton = document.getElementById(cId + '_atrImg');
	if (allToRightButton != null){
		allToRightButton.disabled = disabled;
	}
	
	allToLeftButton = document.getElementById(cId + '_atlImg');
	if (allToLeftButton != null){
		allToLeftButton.disabled = disabled;	
	}	
}

function SelectByText(listId,text){
	myList = document.getElementById(listId);
	currentSelIndex = myList.selectedIndex;	
	var i = 0;
	var index = -1;
	while(i<myList.length){
		itemText = myList.options[i].text;
		if(itemText.toLowerCase().substring(0,text.length)==text){
			if(index==-1){
				myList.options[i].selected=true;
			}
			else{
				myList.options[i].selected=false;
			}
			index = i;						
		}
		else{
			myList.options[i].selected=false;
		}
		i++;
	}
	
	if(index==-1){
		if(currentSelIndex < 0){
			myList.options[0].selected=true;
		}
		else{
			myList.options[currentSelIndex].selected=true;
		}		
	}			
}


function all2right(lListId,rListId,hiddenId)
{

	hidden = document.getElementById(hiddenId);			
	lList =  document.getElementById(lListId);
	rList = document.getElementById(rListId);
	lListlen = lList.length ;
	
	//only if not disabled
	if (lList.disabled==false){
		//Add items in the right list
		for (i=0; i<lListlen ; i++)
		{
			rListlen = rList.length;
			rList.options[rListlen]= new Option(lList.options[i].text);
			rList.options[rListlen].value = lList.options[i].value;					
			if (hidden.value=='')
			{
				hidden.value += lList.options[i].value;
			}			
			else
			{
				hidden.value += '#' + lList.options[i].value;
			}
		}
	
	
		//Remove all items from the left list
		for (i=(lListlen-1); i>=0; i--) 
		{	
			lList.options[i] = null;
		}
	}

}

function right2left(lListId,rListId,hiddenId)
{
	hidden = document.getElementById(hiddenId);	
	lList =  document.getElementById(lListId);				
	rList = document.getElementById(rListId);
	rListLen = rList.length;
	lListLen = lList.length;
	
	for (i=0; i<rListLen ; i++)
	{						
		if (rList.options[i].selected == true) 
		{
			lListlen = lList.length;
			index = GetInsertPlaceForListItem(lList,rList.options[i].text);
			newOption = new Option(rList.options[i].text);
			newOption.value = rList.options[i].value;			
			lList.options.add(newOption,index);			
		}
	}		
	
	hidden.value = '';
	for (i = (rListLen -1); i>=0; i--){
		if (rList.options[i].selected == true) 
		{
			rList.options[i] = null;
		}
		else
		{
			if (hidden.value=='')
			{
				hidden.value += rList.options[i].value;
			}			
			else
			{
				hidden.value += '#' + rList.options[i].value;
			}	
		}	
	}	
}

function left2right(lListId,rListId,hiddenId) 
{								
	hidden = document.getElementById(hiddenId);	
	lList =  document.getElementById(lListId);				
	rList = document.getElementById(rListId);
	lListlen = lList.length ;
	
	for (i=0; i<lListlen ; i++)
	{						
		if (lList.options[i].selected == true) 
		{
			rListlen = rList.length;
			index = GetInsertPlaceForListItem(rList,lList.options[i].text);
			newOption = new Option(lList.options[i].text);
			newOption.value = lList.options[i].value;			
			rList.options.add(newOption,index);				
			if (hidden.value=='')
			{
				hidden.value += lList.options[i].value;
			}			
			else
			{
				hidden.value += '#' + lList.options[i].value;
			}			
		}
	}

	for (i = (lListlen -1); i>=0; i--){
		if (lList.options[i].selected == true) 
		{
			lList.options[i] = null;
		}
	}								
}

function GetInsertPlaceForListItem(list,itemText){
	var i = 0;
	var index = -1;
	while(index==-1 && i<list.length){		
		if (itemText.toLowerCase() < list.options[i].text.toLowerCase()){			
			index=i;
		}		
		i++;
	}		
	
	if (index==-1){
		index = list.length;
	}
	return index;
}

function all2left(lListId,rListId,hiddenId)
{
	hidden = document.getElementById(hiddenId);			
	lList =  document.getElementById(lListId);
	rList = document.getElementById(rListId);
	rListlen = rList.length ;
	
	//only if not disabled
	if (rList.disabled==false){
		//Add items to the left list
		for (i=0; i<rListlen ; i++)
		{
			lListlen = lList.length;
			lList.options[lListlen]= new Option(rList.options[i].text);
			lList.options[lListlen].value = rList.options[i].value;			
		}
		
		hidden.value='';
	
		//Remove all items from the right list
		for (i=(rListlen-1); i>=0; i--) 
		{	
			rList.options[i] = null;
		}
	}
}