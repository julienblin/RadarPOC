oItemStyle = new itemStyle();
oHover = new itemStyle();
oHover.backgroundColor = '#C3A794'

function itemStyle(){
	this.backgroundColor = "";
}

function InitDropDownButtons()
{
	for(i=0; i<buttonIds.length; i++)
	{
		var dropDownId = buttonIds[i].substr(0,buttonIds[i].indexOf("_div"));
		if(document.all(dropDownId) != null)
		{
			document.all(dropDownId + "_hid").value = "";
			var beginTag = "<tr class='";
			var beginTagCss = "ddButtonRow'><td onclick=btnSbmit('";
			var beginTagACss = "ddButtonAlternateRow'><td onclick=btnSbmit('";
			var middleTag = "'); onmouseout='btnMseOut();' onmouseover='btnMseOver();' class='ddButtonCell'><nobr>";
			var endTag = "</nobr></td></tr>"
			var itemsField = document.all(dropDownId + "_itemsField");
			if(itemsField != null && itemsField.value != null && itemsField.value != "")
			{
				var itemList = document.all(dropDownId + "_itemList");
				var reCss = new RegExp("__","gi")
				var cssItems = itemsField.value.replace(reCss, beginTagACss);
				var reCssA = new RegExp("#","gi")
				var allItems = cssItems.replace(reCssA, beginTagCss);
				var reCodes = new RegExp("@", "gi");
				var items = allItems.replace(reCodes, middleTag);
				var reItems = new RegExp("~~", "gi");
				var items = beginTag + items.replace(reItems, endTag + beginTag) + endTag;
				var reTableStart = new RegExp(".*<tbody>", "gi")
				var reTableEnd = new RegExp("</tbody>.*", "gi")
				var tableStart = reTableStart.exec(itemList.outerHTML);
				var tableEnd = reTableEnd.exec(itemList.outerHTML);
				itemList.outerHTML = tableStart + items + tableEnd;
			}
		}
	}
}

function ShowHideButtonDiv(sect){

	var MyDiv = document.getElementById(sect);
	if(MyDiv.style.display=='none')
	{	
		MyDiv.style.display = '';
		var image = document.all(sect.replace("_div", "_image"));
		var currentHeight = parseInt(MyDiv.clientHeight);
		var maxHeight = parseInt(MyDiv.getAttribute("maxHeight"))
		if(currentHeight > maxHeight)
			{
				MyDiv.style.height = parseInt(MyDiv.getAttribute("maxHeight"));
			}
		//Hack for IE scrollbar:
		MyDiv.style.width = parseInt(MyDiv.clientWidth);
	  	MyDiv.style.marginTop = image.height - 4;
	  	MyDiv.style.marginLeft = -parseInt(MyDiv.style.width) - 6;

	}
	else
		MyDiv.style.display='none';
	
	dontHideDropDownButton = 1;
	HideOtherDiv(sect)
}

function btnMseOut(){
	myTd = window.event.srcElement;
	if (myTd.tagName.toLowerCase()!='td'){
		myTd = myTd.parentElement;		
	}
	myTd.style.backgroundColor=oItemStyle.backgroundColor;
}

function btnMseOver(){	
	myTd = window.event.srcElement;
	if (myTd.tagName.toLowerCase()!='td'){
		myTd = myTd.parentElement;		
	}		
	// Set background color property
	oItemStyle.backgroundColor = myTd.style.backgroundColor;
	myTd.style.backgroundColor= oHover.backgroundColor;
}

function btnSbmit(code){
	var codeArr = code.split("~");
	if(codeArr[1].toLowerCase() == "true" && this.initProgress)
		initProgress();
	mySpan = btnGetMainSpan(window.event.srcElement);
	myHidden = document.getElementById(mySpan.getAttribute('hidId'))
	myHidden.value = codeArr[0];
	document.forms.item(0).fireEvent("onsubmit");
	document.forms.item(0).submit();
}

function btnGetMainSpan(td){
	var curObj = td
	while (curObj != window && curObj.getAttribute('hidId') == null){
		curObj = curObj.parentElement;
	}
	return curObj;
}
