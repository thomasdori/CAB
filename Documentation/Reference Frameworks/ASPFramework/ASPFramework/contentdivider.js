if (window.addEventListener || window.attachEvent){
document.write('<style type="text/css">\n')
document.write('.dropcontent{display:none;}\n')
document.write('</style>\n')
}

// Content Tabs script- By JavaScriptKit.com (http://www.javascriptkit.com)
// Last updated: July 25th, 05'

var showrecords=1 //specify number of contents to show per tab
var tabhighlightcolor="yellow" //specify tab color when selected
var taboriginalcolor="lightyellow" //specify default tab color. Should echo your CSS file definition.

////Stop editing here//////////////////////////////////////
document.getElementsByClass=function(tag, classname){
var tagcollect=document.all? document.all.tags(tag): document.getElementsByTagName(tag) //IE5 workaround
var filteredcollect=new Array()
var inc=0
for (i=0;i<tagcollect.length;i++){
if (tagcollect[i].className==classname)
filteredcollect[inc++]=tagcollect[i]
}
return filteredcollect
}


function contractall(){
var inc=0
while (contentcollect[inc]){
contentcollect[inc].style.display="none"
inc++
}
}

function expandone(whichpage){
var lowerbound=(whichpage-1)*showrecords
var upperbound=(tabstocreate==whichpage)? contentcollect.length-1 : lowerbound+showrecords-1
contractall()
for (i=lowerbound;i<=upperbound;i++)
contentcollect[i].style.display="block"
}

function highlightone(whichtab){
for (i=0;i<tabscollect.length;i++){
tabscollect[i].style.backgroundColor=taboriginalcolor
tabscollect[i].style.borderRightColor="white"
tabsfootcollect[i].style.backgroundColor="white"
}
tabscollect[whichtab].style.backgroundColor=tabhighlightcolor
tabscollect[whichtab].style.borderRightColor="gray"
tabsfootcollect[whichtab].style.backgroundColor="#E8E8E8"
}

function generatetab(){
contentcollect=document.getElementsByClass("p", "dropcontent")
tabstocreate=Math.ceil(contentcollect.length/showrecords)
linkshtml=""
linkshtml2=""
for (i=1;i<=tabstocreate;i++){
linkshtml+='<span class="tabstyle" onClick="expandone('+i+');highlightone('+(i-1)+')"><b>Page '+i+'</b></span>'
linkshtml2+='<a href="#" class="tabsfootstyle" onClick="expandone('+i+');highlightone('+(i-1)+');return false">Page '+i+'</a> '
}
document.getElementById("cyclelinks").innerHTML=linkshtml
document.getElementById("cyclelinks2").innerHTML=linkshtml2
tabscollect=document.getElementsByClass("span", "tabstyle")
tabsfootcollect=document.getElementsByClass("a", "tabsfootstyle")
highlightone(0)
expandone(1)
}

if (window.addEventListener)
window.addEventListener("load", generatetab, false)
else if (window.attachEvent)
window.attachEvent("onload", generatetab)