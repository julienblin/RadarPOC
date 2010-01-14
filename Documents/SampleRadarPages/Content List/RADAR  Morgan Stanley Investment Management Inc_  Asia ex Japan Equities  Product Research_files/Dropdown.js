function InitDropDowns()
{
	for(i=0; i<ids.length; i++)
	{
		var dropDownId = ids[i].substr(0,ids[i].indexOf("_div"));
		var dropDown = document.all(dropDownId);
		if(dropDown != null)
		{
			var myHidden = document.getElementById(dropDown.getAttribute('hidId'));
			myHidden.value = "";
			var itemsField = document.all(dropDownId + "_itemsField");
			var itemList = document.all(dropDownId + "_itemList");			
			var beginTag = "<tr class='ddRow'><td onclick='sbmit();' onmouseout='mseOut();' onmouseover='mseOver();' class='ddCell'><nobr>"
			var endTag = "</nobr></td></tr>"
			var reItems = new RegExp("~~", "gi");			
			var items = beginTag + itemsField.value.replace(reItems, endTag + beginTag) + endTag;
			
			var reTableStart = new RegExp(".*<tbody>", "gi")
			var reTableEnd = new RegExp("</tbody>.*", "gi")
			var tableStart = reTableStart.exec(itemList.outerHTML);
			var tableEnd = reTableEnd.exec(itemList.outerHTML);
			
			itemList.outerHTML = tableStart + items + tableEnd;
			
			//var lblChoice = document.getElementById(dropDownId + "_lblChoice");
			//var labelDiv = lblChoice.parentElement;
			var maxWidth = dropDown.getAttribute("maxWidth");
			if(maxWidth != null && maxWidth != "" && maxWidth.indexOf("%") == -1)
				dropDown.style.width = Math.min(dropDown.offsetWidth, parseInt(maxWidth));
			/*if(dropDown.getAttribute("maxWidth") != null && dropDown.getAttribute("maxWidth") != "")
				width = Math.max(width, parseInt(dropDown.getAttribute("maxWidth")))*/
			//labelDiv.style.width = width;
		}
	}
}

function ShowHideDiv(sect){	

	var MyDiv = document.getElementById(sect);
	if(MyDiv.style.display=='none')
	{
		MyDiv.style.display='';
	  	var currentHeight = parseInt(MyDiv.clientHeight);
	  	var maxHeight = parseInt(MyDiv.getAttribute("maxHeight"))
		if(currentHeight > maxHeight)
			MyDiv.style.height = MyDiv.getAttribute("maxHeight");
		MyDiv.style.width = MyDiv.clientWidth;
		openShim(MyDiv);
	}
	else
	{
		MyDiv.style.display='none';
		closeShim(MyDiv);
	}
		
	dontHideDropDown = 1;
	HideOtherDiv(sect)
}

function mseOut(){
	myTd = window.event.srcElement;
	if (myTd.tagName.toLowerCase()!='td'){
		myTd = myTd.parentElement;		
	}
	myTd.style.backgroundColor='#ffffff';	
}

function mseOver(){	
	myTd = window.event.srcElement;
	if (myTd.tagName.toLowerCase()!='td'){
		myTd = myTd.parentElement;		
	}		
	myTd.style.backgroundColor='#DFCEBC';	
}

function sbmit(){	
	mySpan = getMainSpan(window.event.srcElement);	
	myHidden = document.getElementById(mySpan.getAttribute('hidId'))	
	myTd = window.event.srcElement;
	if(myHidden.value != myTd.innerText)
	{		
		myHidden.value = myTd.innerText;
		document.forms.item(0).submit();
	}	
}

function getMainSpan(td){
	var getOut=0;
	var count;
	var curObj = td
	while (getOut==0){
		tag = curObj.tagName.toLowerCase();
		if (tag =='span'){			
			if (curObj.getAttribute('hidId')!=''){				
				return curObj;	
			}			
		}
		curObj = curObj.parentElement;
		count++;
		if (count==500){			
			getOut=1;
		}
	}
}
