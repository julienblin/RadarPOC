function TabOnMouseOver(sender, tabStripId)
{
	var tabStrip = document.all(tabStripId);
	var hoverClass = tabStrip.getAttribute("hoverClass");
	var defaultClass = tabStrip.getAttribute("defaultClass");
	var leftImageHoverUrl = tabStrip.getAttribute("LeftImageHover");
	var rightImageHoverUrl = tabStrip.getAttribute("RightImageHover");
	
	var leftImage = document.all(sender.getAttribute("leftImageId"));
	var rightImage = document.all(sender.getAttribute("rightImageId"));
	
	//If the tab is not selected and not in hover state, set it hover
	if(sender.className == defaultClass)
	{
		sender.className = hoverClass;
		leftImage.src = leftImageHoverUrl;
		rightImage.src = rightImageHoverUrl;
	}
}

function TabOnMouseOut(sender, tabStripId)
{
	var tabStrip = document.all(tabStripId);
	var hoverClass = tabStrip.getAttribute("hoverClass");
	var defaultClass = tabStrip.getAttribute("defaultClass");
	var leftImageUrl = tabStrip.getAttribute("LeftImage");
	var rightImageUrl = tabStrip.getAttribute("RightImage");
	
	var leftImage = document.all(sender.getAttribute("leftImageId"));
	var rightImage = document.all(sender.getAttribute("rightImageId"));
	
	//If the tab is in hover state, set it to default state
	if(sender.className == hoverClass)
	{
		sender.className = defaultClass;
		leftImage.src = leftImageUrl;
		rightImage.src = rightImageUrl;
	}
}

function TabOnClick(sender, tabStripId, index)
{
	var tabStrip = document.all(tabStripId);
	var selectedClass = tabStrip.getAttribute("selectedClass");
	var defaultClass = tabStrip.getAttribute("defaultClass");
	
	//Ignore the click if the tab is already selected
	if(sender.className != selectedClass)
	{
		var clientSideCommand = sender.getAttribute("clientSideCommand");
		var autoPostBack = (tabStrip.getAttribute("autoPostBack").toLowerCase() == "true");
		var selectedIndex = tabStrip.getAttribute("selectedIndex");
		
		var selectedTab = document.all(buildTabId(tabStripId, selectedIndex));
		var newTab = sender;
		
		var leftImageUrl = tabStrip.getAttribute("LeftImage");
		var rightImageUrl = tabStrip.getAttribute("RightImage");
		var leftImageSelectedUrl = tabStrip.getAttribute("LeftImageSelected");
		var rightImageSelectedUrl = tabStrip.getAttribute("RightImageSelected");
		
		var selectedLeftImage = document.all(selectedTab.getAttribute("leftImageId"));
		var selectedRightImage = document.all(selectedTab.getAttribute("rightImageId"));
		var newLeftImage = document.all(newTab.getAttribute("leftImageId"));
		var newRightImage = document.all(newTab.getAttribute("rightImageId"));
		
		//Change the state of the previously selected tab and the new selected tab
		selectedTab.className = defaultClass;
		newTab.className = selectedClass;
		selectedLeftImage.src = leftImageUrl;
		selectedRightImage.src = rightImageUrl;
		newLeftImage.src = leftImageSelectedUrl;
		newRightImage.src = rightImageSelectedUrl;
		
		if(autoPostBack)
		{
			if(clientSideCommand && clientSideCommand != "")
				eval(clientSideCommand);
			
			//If AutoPostBack is set to true, simply post the new index to the server
			__doPostBack(ReplaceAll(tabStrip.getAttribute('tabStripId'), '_', '$'), index);
		}
		else
		{
			//Set the previously visible panel invisible and the make the panel
			//associated to the selected tab visible
			var contentHolderId = tabStrip.getAttribute("contentHolderId");			
			var selectedPage = getContentHolder(contentHolderId, selectedIndex);			
			var newPage = getContentHolder(contentHolderId, index);			
			
			tabStrip.setAttribute("selectedIndex", index);
			
			selectedPage.style.display = "none";
			newPage.style.display = "";
			
			if(clientSideCommand && clientSideCommand != "")
				eval(clientSideCommand);
		}
	}
}

function ReplaceAll(value, match, replacement)
{
	var newValue = value;
	while(newValue.match(match) != null)
	{
		newValue = newValue.replace(match, replacement);
	}
	return newValue;
}

function getContentHolder(contentHolderId, index)
{
	var contentHolder = document.all(contentHolderId);
	return contentHolder.childNodes(index);
}

function buildTabId(tabStripId, index)
{
	return tabStripId + "_T" + index;
}