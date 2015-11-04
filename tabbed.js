//Available tabs
var tabs = new Array("env","erg","pla","ele","chm","bio","rad");
var tabs2 = new Array("elim","admin","ppa","emer");
var hideTabs; 
var hideTabsc;

function change(oldClass, newClass) 
{
   var tagged=document.getElementsByTagName('a');
   for(var i = 0 ; i < tagged.length ; i++)
   {
      if (tagged[i].className==oldClass)
      {
         tagged[i].className=newClass;
      }
   }
}

function displayTab(tab)
{
   for (var i=0; i < 7; i++)
   {
     document.getElementById("group" + i).style.display = "none";
   }
   
   if (tab != -1)
   {
      document.getElementById("group" + tab).style.display = "block";
   }
}

function displayTabControl(tab)
{
   for (var i=0; i < 4; i++)
   {
     document.getElementById("groupc" + i).style.display = "none";
   }
   
   if (tab != -1)
   {
      document.getElementById("groupc" + tab).style.display = "block";
   }
}

function switchTab(tabNum)
{
   for (var i=0; i<tabs.length; i++)
   {
      //We just want to apply the CSS on start-up
      //Clear all the Ids if any exists      
      if (document.getElementById("tablink-" + tabs[i]))
         document.getElementById("tablink-" + tabs[i]).id = "";
         
      if (document.getElementById("group-" + tabs[i]))
         document.getElementById("group-" + tabs[i]).id = "";
         
      change(tabs[i] + "on", tabs[i]);       
   }
   
   if (tabNum != -1)
   {
      change(tabs[tabNum], tabs[tabNum] + "on");       
   }
   
   displayTab(tabNum);
   
   //Clear any timeout of hiding tabs
   clearTimeout(hideTabs);
}

function switchTabControl(tabNum)
{
   for (var i=0; i<tabs2.length; i++)
   {
      //We just want to apply the CSS on start-up
      //Clear all the Ids if any exists      
      if (document.getElementById("tablinkc-" + tabs2[i]))
         document.getElementById("tablinkc-" + tabs2[i]).id = "";
         
      if (document.getElementById("groupc-" + tabs2[i]))
         document.getElementById("groupc-" + tabs2[i]).id = "";
         
      change(tabs2[i] + "on", tabs2[i]);       
   }
   
   if (tabNum != -1)
   {
      change(tabs2[tabNum], tabs2[tabNum] + "on");       
   }
   
   displayTabControl(tabNum);
   
   //Clear any timeout of hiding tabs
   clearTimeout(hideTabsc);
}

//Hide all tabs and link groups
function hideAllTabs()
{
   hideTabs = setTimeout("switchTab(-1)",10000);
}

function hideAllTabsControl()
{
   hideTabsc = setTimeout("switchTabControl(-1)",10000);
}