// Javascript Graph Library
// Copyright © 1998 Netscape Communications Corporation. All Rights Reserved.

function Graph(width, height, stacked) {
	this.stacked = stacked;
	this.width = width || 400;
	this.height = height || 200;
	this.rows = new Array();
	this.addRow = _addRowGraph;
	this.setXScale = _setXScaleGraph;
	this.setXScaleValues = _setXScaleValuesGraph;
	this.setTime = _setStartTimeGraph;
	this.setDate = _setStartDateGraph;
	this.build = _buildGraph;
	this.setLegend = _setLegendGraph;
	this.writeLegend = _writeLegendGraph;
	this.offset = 0;
	this.imageStyle = '';
	this.graphStyle ="";
	return this;
};

function _setLegendGraph() {
	this.legends = arguments;
};

function _addRowGraph() {
	this.rows[this.rows.length] = new Array();
	var row = this.rows[this.rows.length-1];
	for(var i = 0; i < arguments.length; i++)
	row[row.length] = arguments[i];
};
function _rescaleGraph(g) {
	g.posMax = 0, g.negMax = 0, g.c = 0;
	for(var i = 0; i < g.rows.length; i++) {
		for(var j = 0; j < g.rows[i].length; j++) {
			g.c++;
			if(g.rows[i][j] > g.posMax) g.posMax = g.rows[i][j];
			if(g.rows[i][j] < g.negMax) g.negMax = g.rows[i][j];
		}
	}
	g.vscale = g.height/(g.posMax-g.negMax);
	g.hscale = Math.floor(g.width/g.c-1/g.rows[0].length);
};

function _stackRescaleGraph(g) {
	var m, n, c = 0;
	g.posMax = 0, g.negMax = 0;
	for(var i = 0; i < g.rows[0].length; i++) {
		m = 0; n = 0;
		c++;
		for(var j = 0; j < g.rows.length; j++) {
			if(g.rows[j][i] > 0) m += g.rows[j][i];
			else n += g.rows[j][i];
		}
		if(m > g.posMax) g.posMax = m;
		if(n < g.negMax) g.negMax = n;
	}
	g.vscale = g.height/(g.posMax-g.negMax);
	g.hscale = Math.floor(g.width/c)-1;
};

function _relRescaleGraph(g) {
	var m, c = 0;
	g.vscale = g.height/100;
	for(var i = 0; i < g.rows[0].length; i++) {
		m = 0;
		c++;
		for(var j = 0; j < g.rows.length; j++) {
			if(g.rows[j][i] < 0) g.rows[j][i] = 0;
			m += g.rows[j][i];
		}
		var s = 100/m; var k = 0;
		for(var j = 1; j < g.rows.length; j++) {
			g.rows[j][i] *= s;
			g.rows[j][i] = Math.round(10*g.rows[j][i])/10;
			if(j != 0) k += g.rows[j][i]*g.vscale;
		}
		g.rows[0][i] = Math.round(10*(g.height-k)/g.vscale)/10;
	}
	g.hscale = Math.floor(g.width/c)-1;
	g.posMax = 100; g.negMax = 0;
};

function _writeLegendGraph() {
var st = "";
st += "<table style='border:1 black solid;' border='0' cellspacing='0' cellpadding='4'><tr><td><font face='arial,helvetica' size='-1'>";
for(var i = 0; i < this.legends.length; i++) {
if(!this.legends[i]) continue;
if(i >= this.rows.length) break;
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/"+gImage(i)+".gif' border='1' hspace='3' width='5' height='5'>"+this.legends[i]+"<BR>\n";
}
st += "</font></td></tr></table>";
return st;
};

function _buildRegGraph(g, doc) {
var str = "";
str += "<table style='"+g.graphStyle+"' border='0' cellpadding='0' cellspacing='0'>\n";
if(g.title) {
str += "<tr>\n";
if(g.scale) str += "<td colspan='3'></td>\n";
if(g.yLabel) str += "<td></td>\n";
str += "<th valign='top' height='30' colspan='"+(g.c)+"'>\n";
str += "<font face='arial,helvetica' size='-1'>";
str += g.title;
str += "</font></th></tr>\n";
}
if(g.yLabel) {
g.yLabel = g.yLabel.split("");
g.yLabel = g.yLabel.join("<BR>");
str += "<tr>\n";
var r = 2; if(g.negMax && g.posMax) r++;
str += "<th rowspan="+r+" align=center width=20 nowrap>\n";
str += "<font face='arial,helvetica' size='-1'>"+g.yLabel+"</font></td>\n";
}
if(g.posMax > 0) {
if(!g.yLabel) str += "<tr>\n";
if(g.scale) str += _writeScaleGraph(g, 0, g.posMax);
str += "<td valign=bottom";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += ">";
for(var j = 0; j < g.rows[0].length; j++) {
for(var i = 0; i < g.rows.length; i++) {
if(parseInt(g.vscale*g.rows[i][j]) > 0) {
str += "<img style='"+g.imageStyle+"' border=0 src='"+SCRIPT_LIBRARY_PATH+"graph/images/"+gImage(i)+".gif' ";
str += "alt=\"";
if(g.legends && g.legends[i]) str += g.legends[i]+": ";
str += (g.rows[i][j]+g.offset);
if(g.dates) str += ", "+g.dates[j];
str += "\" ";
str += "width="+parseInt(g.hscale)+" ";
str += "height="+parseInt(g.vscale*g.rows[i][j])+" ";
str += ">";
} else
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width="+parseInt(g.hscale)+" height='5'>";
str += "</td>\n<td valign='bottom'";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += ">";
}
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='5'>";
}
str += "</td>\n";
}
if(g.legends && g.posMax != 0) {
str += "<td width=5 nowrap rowspan=3></td>\n";
str += "<td rowspan=3>";
str += g.writeLegend();
str += "</td>\n";
}
if(g.scale || g.xScale) {
if(g.posMax) str += "</tr><tr>\n";
else str += "<tr><td colspan='2'></td>\n";
str += "<td bgcolor='#000000' colspan='"+(g.c+1)+"'>";
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' height=1 width=";
str += parseInt(g.rows[0].length*g.hscale)+" ></td></tr>\n";
}
if(g.xScale && !g.negMax)
str += _writeXScaleGraph(g);
if(g.negMax < 0) {
if(g.posMax != 0 && !g.scale) str += "</tr>";
str += "<tr>\n";
if(g.scale) str += _writeNegScaleGraph(g, g.negMax, 0);
str += "<td VALIGN=TOP";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += ">";
for(var j = 0; j < g.rows[0].length; j++) {
for(var i = 0; i < g.rows.length; i++) {
if(parseInt(g.vscale*g.rows[i][j]) < 0) {
str += "<img  vspace='0' hspace='0' border='0' align='top' src='"+SCRIPT_LIBRARY_PATH+"graph/images/"+gImage(i)+".gif' width='"+
parseInt(g.hscale)+"' height='"+
parseInt(-1*g.vscale*g.rows[i][j]);
str += "' alt=\"";
if(g.legends && g.legends[i]) str += g.legends[i]+": ";
str += (g.rows[i][j]+g.offset);
if(g.dates) str += ", "+g.dates[j];
str += "\" >";
} else
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' align='top' border='0' width='"+parseInt(g.hscale)+
"' height='5'>";
str += "</td>\n<td valign='top'";
if(g.bgColor) str += " bgcolor='"+g.bgColor+"'";
str += ">";
}
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' ALIGN=TOP width=1 BORDER=0 height=5>";
}
str += "</td>\n";
if(g.legends && g.posMax == 0) {
str += "<td width=5 NOWRAP ROWSPAN=2></td><td>";
str += g.writeLegend();
str += "</td>\n";
}
}
str += "</td></tr>\n";
if(g.xLabel) {
str += "<tr>\n";
if(g.scale) str += "<td colspan=3></td>\n";
if(g.yLabel) str += "<td></td>\n";
str += "<TH colspan="+g.c+" height=25 VALIGN=BOTTOM><FONT FACE='Arial,Helvetica' SIZE=-1>";
str += g.xLabel;
str += "</font></th></tr>\n";
}
str += "</table>\n";
doc.write(str);
};

function _setXScaleGraph(s, skip, inc) {
this.xScale = true;
this.s = s || 0;
this.skip = skip || 1;
this.inc = inc || 1;
};

function _setXScaleValuesGraph() {
this.xScale = true;
this.s = 0;
this.skip = 1;
this.inc = 1;
this.dates = new Array();
for(var i = 0; i < arguments.length; i++)
this.dates[this.dates.length] = arguments[i];
};

function _setStartTimeGraph(hour, min, skip, inc) {
this.xScale = true;
this.sTime = new Date(0, 0, 0, hour, min);
this.skip = skip || 1;
this.inc = inc || 1;
};

function _setStartDateGraph(month, day, year, skip, inc) {
this.xScale = true;
this.sDate = new Date(year, month-1, day);
this.skip = skip || 1;
this.inc = inc || skip || 1;
this.showDate = true;
};

function _setDatesArrayGraph(g) {
if(g.dates) return;
g.dates = new Array();
for(var i = 0; i < g.rows[0].length; i++) {
var dateStr = "";
if(g.sDate) {
if(g.showDay) {
eval('switch(g.sDate.getDay()) {'+
'case 0: dateStr += "Sun"; break;'+
'case 1: dateStr += "Mon"; break;'+
'case 2: dateStr += "Tue"; break;'+
'case 3: dateStr += "Wed"; break;'+
'case 4: dateStr += "Thu"; break;'+
'case 5: dateStr += "Fri"; break;'+
'case 6: dateStr += "Sat"; break;'+
'}');
dateStr += " ";
}
if(g.longDate && g.showDate) {
dateStr += g.sDate.getDate()+"-";
eval('switch(g.sDate.getMonth()) {'+
'case 0: dateStr += "Jan"; break;'+
'case 1: dateStr += "Feb"; break;'+
'case 2: dateStr += "Mar"; break;'+
'case 3: dateStr += "Apr"; break;'+
'case 4: dateStr += "May"; break;'+
'case 5: dateStr += "Jun"; break;'+
'case 6: dateStr += "Jul"; break;'+
'case 7: dateStr += "Aug"; break;'+
'case 8: dateStr += "Sep"; break;'+
'case 9: dateStr += "Oct"; break;'+
'case 10: dateStr += "Nov"; break;'+
'case 11: dateStr += "Dec"; break;'+
'}');
} else if(g.showDate) dateStr += (g.sDate.getMonth()+1)+"/"+g.sDate.getDate();
if(g.showYear && g.showDate) {
if(g.longDate) dateStr += "-";
else dateStr += "/";
}
if(g.showYear) {
if(g.longYear) dateStr += g.sDate.getFullYear();
else dateStr += (g.sDate.getFullYear()%100);
}
g.sDate.setDate(g.sDate.getDate()+ g.inc);
} else if(g.sTime) {
var hrs = g.sTime.getHours();
if(!g.armyTime) {
var pm = false;
if(hrs == 0) { hrs = 12; }
else if(hrs >= 12) { if(hrs > 12) hrs -= 12; pm = true; }
} else
if(hrs < 10) hrs = "0" + hrs;
dateStr = hrs + ":";
var min = g.sTime.getMinutes();
if(min < 10) min = "0" + min;
dateStr += min;
if(!g.armyTime) { !pm ? dateStr += "am" : dateStr += "pm" ; }
g.sTime.setMinutes(g.sTime.getMinutes()+ g.inc);
} else dateStr = g.s+i*g.inc;
g.dates[i] = dateStr;
}
};

function _writeXScaleGraph(g) {
var st = "";
if(!g.c) g.c = g.rows[0].length*2-1;
st += "<tr>\n";
if(g.scale) st += "<td colspan=2></td>\n";
if(g.yLabel) st += "<td></td>\n";
st += "<td VALIGN=TOP colspan="+(g.c+1)+">";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' height='10' width='1'>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' height='1' width='1'>";
var n = g.rows[0].length;
var mult = g.rows.length;
if(g.stacked || g.relative) mult = 1;
for(var i = 0; i < n; i++) {
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' height='1' width="+(g.hscale*mult)+">";
if((i+1) % g.skip)
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' height='10' width='1'>";
else
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' height='10' width='1'>";
}
st += "</td></tr><tr>\n";
if(g.scale) st += "<td colspan=3></td>\n";
if(g.yLabel) st += "<td></td>\n";
var cspan = g.rows.length;
if(g.stacked || g.relative) cspan = 2;
cspan *= g.skip;
if(g.sDate || g.sTime) _setDatesArrayGraph(g);
var t = 0;
for(var i = 0; i < Math.floor(g.rows[0].length/g.skip); i++) {
st += "<td VALIGN=TOP";
st += " colspan="+cspan; t += cspan;
st += "><FONT FACE='Arial,Helvetica' SIZE=-3><I>";
st += g.dates[i*g.skip] || "";
st += "</I></font></td>\n";
}
var len = g.rows[0].length;
if(!g.stacked) len *= g.rows.length;
if(i < Math.ceil(g.rows[0].length/g.skip)) {
st += "<td VALIGN=TOP";
st += " colspan="+(len-t);
st += "><FONT FACE='Arial,Helvetica' SIZE=-3><I>";
st += g.dates[i*g.skip];
st += "</I></font></td>\n";
}
st += "</tr>\n";
return st;
};

function _writeNegScaleGraph(g, min, max) {
var h = Math.ceil(g.height/(g.posMax-g.negMax)*g.scale);
var p = -1*g.negMax/(g.posMax-g.negMax);
var n = Math.floor(g.height*p/h);
var st = "";
if(h < 15) {
if(!g.posMax)
//alert("Warning! Scale is too small! Please make\nthe scale larger or make the graph taller.");
st += "<td></td><td></td><td></td>\n"
return st;
}
st += "<td VALIGN=TOP ALIGN=RIGHT>";
var H = h - 3;
for(var i = 0; i < n; i++) {
st += "<FONT FACE=Arial,Helvetica SIZE=-3><I>"+(g.scale*-1*(i+1)+g.offset)+"</I></font>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+H+"'><BR>\n";
}
st += "</td>\n";
st += "<td VALIGN=TOP>";
for(var i = 0; i < n; i++) {
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+(h-1)+"'><BR>\n";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' width='6' height='1'><BR>\n";
}
st += "</td>\n";
st += "<td VALIGN=TOP>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' width='1' height='"+(g.height*p)+"'>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+(g.height*p)+"'>";
st += "</td>\n"
return st;
};

function _writeScaleGraph(g, min, max) {
var h;
var p = g.posMax/(g.posMax-g.negMax);
h = Math.ceil(g.height/(g.posMax-g.negMax)*g.scale);
var n = Math.floor(g.height*p/h);
var st = "";
if(h < 15) {
//alert("Warning! Scale is too small! Please make\nthe scale larger or make the graph taller.");
st += "<td ROWSPAN=2></td><td ROWSPAN=2></td><td></td>\n"
return st;
}
st += "<td VALIGN=BOTTOM ROWSPAN=2 ALIGN=RIGHT>";
var H = h - 3;
for(var i = 0; i < n; i++) {
st += "<FONT FACE=Arial,Helvetica SIZE=-3><I>"+(g.scale*(n-1)-g.scale*i+g.offset);
if(g.relative) st += "%";
st += "</I></font>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+H+"'><BR>\n";
}
st += "</td>\n";
st += "<td valign=bottom rowspan=2>";
for(var i = 0; i < n; i++) {
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' width='6' height='1'><BR>\n";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+(h-1)+"'><BR>\n";
}
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' width='6' height='1'><BR>\n";
st += "</td>\n";
st += "<td valign=bottom>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' width='1' height='"+(g.height*p)+"'>";
st += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='"+(g.height*p)+"'>";
st += "</td>\n"
return st;
};

function _buildStackGraph(g) {
if(!g.c) g.c = g.rows[0].length*2-1;
var str = "";
str += "<table style='"+g.graphStyle+"' border='0' cellpadding='0' cellspacing='0'>\n";
if(g.title) {
str += "<tr>\n";
if(g.scale) str += "<td colspan='3'></td>\n";
if(g.yLabel) str += "<td></td>\n";
str += "<th valign='top' height='30' colspan='"+(g.c)+"'>";
str += "<font face='arial,helvetica' size='-1'>";
str += g.title;
str += "</font></th></tr>\n";
}
if(g.yLabel) {
g.yLabel = g.yLabel.split("");
g.yLabel = g.yLabel.join("<BR>\n");
str += "<tr>\n";
var rspan = 2; if(g.negMax && g.posMax) rspan++;
str += "<TH ROWSPAN="+rspan+" ALIGN=LEFT NOWRAP width=20>";
str += "<FONT FACE='Arial,Helvetica' SIZE=-1>"+g.yLabel+"</font></th>\n";
}
if(g.posMax > 0) {
if(!g.yLabel) str += "<tr>\n";
if(g.scale) str += _writeScaleGraph(g, 0, g.posMax);
for(var j = 0; j < g.rows[0].length; j++) {
str += "<td VALIGN=BOTTOM";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += ">";
var k = 0, y = 0, drawn = false;
for(var i = 1; i < g.rows.length; i++)
if(parseInt(g.vscale*g.rows[i][j]) > 0)
k += parseInt(g.vscale*g.rows[i][j]);
if(g.rows.length > 0 && g.relative && (g.height > k)) {
str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/0.gif' width="+parseInt(g.hscale)+" height="+(g.height-k)+" ";
str += "ALT=\"";
if(g.legends && g.legends[0]) str += g.legends[0]+": ";
str += Math.round((g.height-k)/g.vscale*10)/10+"%";
if(g.dates) str += ", "+g.dates[j];
str += "\" ><BR>\n";
y++;
drawn = true;
}
for(var i = y; i < g.rows.length; i++) {
if(parseInt(g.vscale*g.rows[i][j]) > 0) {
str += "<img  src='"+SCRIPT_LIBRARY_PATH+"graph/images/"+gImage(i)+".gif' width="+parseInt(g.hscale)+" height="+
parseInt(g.vscale*g.rows[i][j])+" ";
str += "ALT=\"";
if(g.legends && g.legends[i]) str += g.legends[i]+": ";
str += g.rows[i][j];
if(g.relative) str += "%";
if(g.dates) str += ", "+g.dates[j];
str += "\" ><BR>\n";
drawn = true;
}
}
if(!drawn) str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width="+parseInt(g.hscale)+" height=1>";
str += "</td>\n";
str += "<td";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += "><img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif width='1' height='5'></td>\n";
}
}
if(g.legends && g.posMax != 0) {
str += "<td width=5 NOWRAP ROWSPAN=3></td>\n<td ROWSPAN=3>";
str += g.writeLegend();
str += "</td>\n";
}
if(g.scale || g.xScale) {
str += "</tr><tr>\n<td bgcolor=#000000 colspan="+(g.rows[0].length*2);
str += "><img src='"+SCRIPT_LIBRARY_PATH+"graph/images/black.gif' height='1' width=";
str += parseInt(g.rows[0].length*g.hscale)+" ></td></tr>\n";
}
if(g.xScale && !g.negMax) {
str += _writeXScaleGraph(g);
}
if(g.negMax < 0) {
if(g.posMax != 0) str += "</tr>\n";
str += "<tr>\n";
if(g.scale) str += _writeNegScaleGraph(g, g.negMax, 0);
for(var j = 0; j < g.rows[0].length; j++) {
str += "<td VALIGN=TOP";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += ">";
var drawn = false;
for(var i = 0; i < g.rows.length; i++) {
if(parseInt(g.vscale*g.rows[i][j]) < 0) {
str += "<img vspace=0 HSPACE=0 BORDER=0 ALIGN=TOP src='"+SCRIPT_LIBRARY_PATH+"graph/images/"+gImage(i)+".gif' width="+
parseInt(g.hscale)+" height="+
parseInt(-1*g.vscale*g.rows[i][j])+" ";
str += "ALT=\"";
if(g.legends && g.legends[i]) str += g.legends[i]+": ";
str += g.rows[i][j];
if(g.relative) str += "%";
if(g.dates) str += ", "+g.dates[j];
str += "\" ><BR>\n";
drawn = true;
}
}
if(!drawn) str += "<img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width="+parseInt(g.hscale)+" height=1>";
str += "</td>\n";
str += "<td";
if(g.bgColor) str += " bgcolor=\""+g.bgColor+"\"";
str += "><img src='"+SCRIPT_LIBRARY_PATH+"graph/images/clear.gif' width='1' height='5'></td>";
}
if(g.legends && g.posMax == 0) {
str += "<td width=5 NOWRAP ROWSPAN=2></td>\n<td>";
str += g.writeLegend();
str += "</td>\n";
}
}
str += "</td></tr>\n";
if(g.xLabel) {
str += "<tr>\n";
if(g.scale) str += "<td colspan=3></td>\n";
if(g.yLabel) str += "<td></td>\n";
str += "<TH colspan="+g.c+" height=25 VALIGN=BOTTOM><FONT FACE='Arial,Helvetica' SIZE=-1>";
str += g.xLabel;
str += "</font></th></tr>\n";
}
str += "</table>\n";
doc.write(str);
};

function _adjustOffsetGraph(g) {
if(g.relative) return;
for(var i = 0; i < g.rows.length; i++)
for(var j = 0; j < g.rows[i].length; j++)
g.rows[i][j] -= g.offset;
};

function _buildGraph(d) {
doc = d || document;
if(!this.rows) return;
if(this.rows.length == 0) {
doc.write("<table><tr><td><tt>[empty graph]</tt></td></tr></table>\n");
return;
}
_adjustOffsetGraph(this);
if(this.xScale) _setDatesArrayGraph(this);
if(this.relative) {
_relRescaleGraph(this);
_buildStackGraph(this, doc);
return;
}
if(this.stacked) {
_stackRescaleGraph(this);
_buildStackGraph(this, doc);
return;
}
_rescaleGraph(this);
_buildRegGraph(this, doc);
};

// Get Graph Image
function gImage(index){
	return index;
	//return Math.floor(Math.random()*9);
};