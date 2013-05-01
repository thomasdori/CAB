/*
	Based on Web-based Content Editor found at Xefteri 
	http://www.xefteri.com/articles/apr202002/default.aspx

*/
function RichTextBox(frame_name,htmlText,readonly,enabled,style,igURL){
	var f;

	this.WebControl = WebControl; // inherit from WebControl
	this.WebControl();
		
	this.name = frame_name;
	this._editMode = 2;  // 1 - html, 2 - wysiwyg
	this._readonly = readonly;
	this._enabled = enabled;
	this._style = style; // document style
	this._igURL = igURL||''; // image gallery url
	this.text = htmlText;
	
	this.editor = f = document.frames[frame_name]; //parent.frames[frame_name];	
	
	this.setText(htmlText);
	window.setTimeout(this+".setDesignMode(true)",100);	
	
};
RichTextBox._cnt = 0;
RichTextBox.Preview = function(text){
	var w=screen.width/1.5,h=screen.height/1.5;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var win = window.open("",'preview_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,scrollbars=yes,resizable=yes,dependent=yes,alwaysRaised=yes');
	if(!win) alert(RichTextBox.__popupBlockerMsg);
	else {
		win.focus();
		win.document.open();
		win.document.write("<html><head><title>Preview</title></head><body>");
		win.document.write(text);
		win.document.write("</body></html>");
		win.document.close();
	}	
};
RichTextBox.ButtonOver = function(style) {
	if(!style) return;
	style.backgroundColor = "#B5BDD6";
	style.borderColor = "darkblue darkblue darkblue darkblue";
	style.borderWidth = '1px';
	style.borderStyle = 'solid'; 
};
RichTextBox.ButtonOut = function(style) {
	if(!style) return;
	style.backgroundColor = "";
	style.borderColor = "";
};
RichTextBox.ButtonDown = function(style) {
	if(!style) return;
	style.backgroundColor = "#8494B5";
	style.borderColor = "darkblue darkblue darkblue darkblue";
};
RichTextBox.ButtonUp = function(style) {
	if(!style) return;
	style.backgroundColor = "#B5BDD6";
	style.borderColor = "darkblue darkblue darkblue darkblue";
};

// Prototype
var p = RichTextBox.prototype = new WebControl;
p._updateValue = function(){
	var f = document.forms[0][this.name+'_Value'];
	if(f) f.value = this.getText();
	window.setTimeout (this+'._updateValue();',400);
};
p.buttonOver = function(eButton)	{
	if (!this._enabled||this._readonly||eButton.leaveOn) return false;	
	RichTextBox.ButtonOver(eButton.style);
};
p.buttonOut = function(eButton) {
	if (!this._enabled||this._readonly||eButton.leaveOn) return false;
	RichTextBox.ButtonOut(eButton.style);
};
p.buttonDown = function(eButton) {
	if (!this._enabled||this._readonly||eButton.leaveOn) return false;	
	RichTextBox.ButtonDown(eButton.style);
};
p.buttonUp = function(eButton) {
	if (!this._enabled||this._readonly||eButton.leaveOn) return false;
	RichTextBox.ButtonUp(eButton.style);
	eButton = null; 
};
p.createLink = function() {
  	if (!this.isDesignMode(true)) return;
  	cmdExec("CreateLink");
};
p.cmdExec = function(cmd,opt) {
  	if (!this.isDesignMode(true)) return;
	this.editor.focus();
  	this.editor.document.execCommand(cmd,"",opt);
};
p.isDesignMode = function(showMsg){
	if (!this._enabled || this._readonly) return false;
	if (this._designMode && this._editMode==2) return true;
	else {
		if(showMsg) alert("Please unselect View HTML");
		return false;
	}
};
p.setDesignMode = function(bMode){	// bMode: true - design, false - view
	if (!this._enabled||this._readonly) bMode=false;
	this._designMode = bMode;
	if(this.editor) {
		this.editor.document.designMode = (bMode)? "On":"Off";
	}
};
p.setEditMode = function(bMode) {	// bMode: 1 - html, 2 - wysiwyg
	var sTmp;
	if (!this._enabled||this._readonly) return false;
  	this._editMode = (bMode==1)? 1:2;
  	if (this._editMode==1) {
		sTmp = this.editor.document.body.innerHTML;
		sTmp = sTmp.replace(/class\=TableOutline/ig,'');
		sTmp = sTmp.replace(/\<style\>\.TableOutline \{ border\:1px dotted silver\;\}\<\/style\>/ig,'');
		this.editor.document.body.innerText=sTmp;
	} else {
		sTmp=this.editor.document.body.innerText;
		sTmp = '<style>.TableOutline { border:1px dotted silver;}</style>' + sTmp;
		sTmp = sTmp.replace(/\<table/ig,"<table class='TableOutline'");
		sTmp = sTmp.replace(/\<td/ig,"<td class='TableOutline'");
		this.editor.document.body.innerHTML=sTmp;
	}
  	this.editor.focus();
};
p.getText = function(t){
	var t;
	if (!this.editor) return '';
	else {
		if(this._editMode==2) t=this.editor.document.body.innerHTML;
		else t=this.editor.document.body.innerText;
		t = t.replace(/class\=TableOutline/ig,'');
		t = t.replace(/\<style\>\.TableOutline \{ border\:1px dotted silver\;\}\<\/style\>/ig,'');
		return t;
	}
};
p.setText = function(t){
	var f = this.editor;
	var doc = (f)? f.document:null;
	this.text = t;
	// table outline style
	t= '<style>.TableOutline { border:1px dotted silver;}</style>' + t;
	t = t.replace(/\<table/ig,"<table class='TableOutline'");
	t = t.replace(/\<td/ig,"<td class='TableOutline'");
	if(doc){
		doc.open();
		doc.write(t);
		doc.close();
	}
	// quick hack to update hidden field
	window.setTimeout (this+'._updateValue();',500);	
};

p.showTableDialog = function() {
	if (!this._enabled||this._readonly||!this.isDesignMode(true)) return false;
	var w=460,h=235;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var url = SCRIPT_LIBRARY_PATH + 'richtext/table.html';
	var win = window.open(url,'table_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,resizable=no,dependent=yes,alwaysRaised=yes');
	if(!win) alert(this.__popupBlockerMsg);
	else {
		win.focus();
		win.CallBackObjectID = this.id;
		win.opener = window;
	}
};
p.createTable = function(arg){
	if (!this._enabled||this._readonly||!this.isDesignMode(true)) return false;
	var w='';
	var table;
	var cursor;
	var trnum=1;
	var tdnum=1;
	if (!arg) return null;
	//----- Creates User Defined Tables -----
	this.editor.focus();
	//width,rows,cols,align,border,padding,spacing
	cursor = this.editor.document.selection.createRange();
	if (!arg.align) arg.align = "left";
	if (!arg.border) arg.border = "0' class='TableOutline";
	if (!arg.rows || arg.rows == 0) arg.rows = 1;
	if (!arg.cols || arg.cols == 0) arg.cols = 1;
	if (arg.tablestyle) arg.tablestyle = "style='"+arg.tablestyle+"'";
	if (arg.tablecolor) arg.tablecolor = "bgcolor='"+arg.tablecolor+"'";
 	if (arg.cellstyle) arg.cellstyle = "style='"+arg.cellstyle+"'";
	if (arg.cellcolor) arg.cellcolor = "bgcolor='"+arg.cellcolor+"'"; 
	if (arg.forecolor) arg.cellstyle = (arg.cellstyle||'')+";color:"+arg.forecolor+";"; 
	table = "<table "+(arg.tablecolor||'')+(arg.tablestyle||'')+" border='"+arg.border+"' align='" + arg.align + "' cellpadding='"+arg.padding+"' cellspacing='"+arg.spacing+"' width='" + arg.width + "'>"
	while (trnum <= arg.rows)	{
		trnum+= 1;
		table+= "<tr>";
		while (tdnum <= arg.cols) {
			if (arg.autosize) w=" width='"+(100/arg.cols)+"%' ";
			table+= "<td valign='top' "+(arg.cellcolor||'')+(arg.cellstyle||'')+w+" class='TableOutline'>&nbsp;</td>";
			tdnum+= tdnum;
		}
		tdnum=1;
		table+= "</tr>";
	}
	table+= "</table>";
	cursor.pasteHTML(table);
};
p.showColorPicker = function(){
	if (!this._enabled||this._readonly||!this.isDesignMode(true)) return false;
	var w=245,h=220;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var url = SCRIPT_LIBRARY_PATH + 'richtext/color.html';
	var win = window.open(url,'color_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,resizable=no,dependent=yes,alwaysRaised=yes');
	if(!win) alert(this.__popupBlockerMsg);
	else {
		win.focus();
		win.CallBackObjectID = this.id;
		win.opener = window;
	}
};
p.setColor = function(c){
	this.cmdExec("ForeColor",c);
};
p.showImageDialog = function() {
	if (!this._enabled||this._readonly||!this.isDesignMode(true)) return false;
	var w=460,h=235;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var ctrl = this.name.split("_")[0];
	var url = SCRIPT_LIBRARY_PATH + 'richtext/image.html';
	
	//get relative url
	var relURL = (window.location.src||window.location.href);
	relURL = relURL.substr(0,relURL.lastIndexOf("/"))+"/";
	var win = window.open(url+"?url="+escape(relURL),'image_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,resizable=no,dependent=yes,alwaysRaised=yes');
	if(!win) alert(this.__popupBlockerMsg);
	else {
		win.focus();
		win.CallBackObjectID = this.id;
		win.opener = window;
	}
};
p.showImageGallery = function(){
	if (!this._enabled||!this.isDesignMode(true)) return false;
	var w=screen.width/2,h=screen.height/2;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var win = window.open(this._igURL,'gallery_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,scrollbars=yes,resizable=yes,dependent=yes,alwaysRaised=yes');
	if(!win) alert(this.__popupBlockerMsg);
	else {
		win.focus();
		win.CallBackObjectID = this.id;
		win.opener = window;
	}	
};
p.setImage = function(url,w,h,b,vs,hs,align){
	var image;
	var cursor;
	// width,height,border,vspace,hspace,align
	image = '<img src="'+url+'"' +
			((w)? ' width ="'+w+'" ':'') +
			((h)? ' height="'+h+'" ':'') +
			((b)? ' border="'+b+'" ':'') +
			((vs)? ' vspace="'+vs+'" ':'') +
			((hs)? ' hspace="'+hs+'" ':'') +
			((align)? ' align ="'+align+'" ':'') +
			' />';
	this.editor.focus();
	cursor = this.editor.document.selection.createRange();
	cursor.pasteHTML(image)	
};
p.toggleHTMLView = function(){
	if (!this._enabled||this._readonly) return false;
	var img = document.images[this.name+'_TBvhtml']; // get toolbar image
	this._HView = !this._HView;
	this.setEditMode(((this._HView)? 1:2));	
	if(img && !this._HView) img.leaveOn = false; 
	else if (img && this._HView) {
		img.leaveOn = true;  // disable hover color change by buttonOver, etc
		img.style.backgroundColor = '#eeeeee';
	}
};
p.showPreview = function(){
	if (!this._enabled||!this.isDesignMode(true)) return false;
	var w=screen.width/1.5,h=screen.height/1.5;
	var l = (screen.width-w)/2;
	var t = (screen.height-h)/2;
	var win = window.open("",'preview_win','top='+ t +',left='+l+',width='+w+',height='+h+',status=yes,scrollbars=yes,resizable=yes,dependent=yes,alwaysRaised=yes');
	if(!win) alert(this.__popupBlockerMsg);
	else {
		win.focus();
		win.CallBackObjectID = this.id;
		win.opener = window;
		win.document.open();
		win.document.write("<html><head><title>Preview</title></head><body>");
		win.document.write(this.getText());
		win.document.write("</body></html>");
		win.document.close();
	}	
};
