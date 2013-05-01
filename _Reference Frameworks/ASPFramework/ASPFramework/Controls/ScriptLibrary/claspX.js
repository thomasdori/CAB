claspCore = function() {
	this.name		= "CLASP";
	this.version	= "1.8";
	this.scriptPath = SCRIPT_LIBRARY_PATH;
	this.events      = new this.notloaded;	
	this.dhtml       = new this.notloaded;
	this.form		 = new claspForm;
	this.ajax        = new claspAjax;
	this.form.validation = new this.notloaded;
	this.page	     = new claspPage;
	this.ie          = (document.all)	   ? true : false;
	this.bwok        = (this.ie);
	this.ns4         = (document.layers) ? true : false;
	this.windowname  = window.name;
};//end clasp core

claspCore.prototype.events      = null;
claspCore.prototype.dhtml       = null;

claspCore.prototype.notloaded = function(	) {
		this.loaded  = false;
		this.version = "1.0";
	}	
	
claspCore.prototype.loadLibrary = function(name) {
	var load = false;
	var libsrc;	
	switch(name) {
		case "dhtml":
			if(!this.dhtml.loaded) {
				load=true;
				libsrc = "dhtml.js";
			}			
			break;
		case "events":
			if(!this.events.loaded) {
				load=true;
				libsrc = "events.js";				
			}
			break;
		case "validation":
			if(!this.form.validation.loaded) {
				load=true;
				libsrc = "validation.js";				
			}
			break;					
	}
	if(load)		
		document.write('<script language="Javascript1.2" src="'+ this.scriptPath + libsrc + '"><\/script>');
}

claspCore.prototype.launchApplication = function(name,width,height) {

	if(window.name == name)	return;
	if(window.opener) if(window.opener.name == name) return;
	
	if(!height) height = screen.height - 100;
	if(!width)  width  = screen.width  - 100;
	var left = (screen.width / 2 ) - width / 2;

	window.name = name + "_close";	
	window.open(location.href + location.search,name,"scrollbars,width=" + width + ",height=" + height + ",top=10,left=" + left)
	clasp.closeAppWindow();					
}

claspCore.prototype.closeAppWindow = function() {
	if(window.history.length == 0){
		self.opener = self;
		self.close()
		return;		
	}	
	if(window.history.length > 0)
		window.history.back();
}

claspCore.prototype.openWindow = function(name,url,left,top,width,height,scrollable,resizable,extraAttributes) {
	var params = "";
	if(left == -1) left =  ((screen.width - width)/2);  
	if(top  == -1)  top =  ((screen.height - height)/2);
	if(scrollable == null) scrollable = false;
	if(resizable  == null) resizable  = false;
	if(extraAttributes=="") extraAttributes = "status=yes,dependent=yes,alwaysRaised=yes"
	params = "left=" + left  + ", " + "top=" + top + ","	
	params = params + "width=" + width + "," + "height=" +  height + "," 
	params = params + "scrollbars=" + (scrollable?"yes":"no") + ","
	params = params + "resizable=" + (resizable?"yes":"no");
	params = params + (extraAttributes?("," + extraAttributes):"")
	return window.open(url,name,params);
}

claspCore.prototype.setSource = function (source) {
	var obj = clasp.getObject('__SOURCE');
	if(obj)
		obj.value = source;
};

claspCore.prototype.setAction = function (action) {
	var obj = clasp.getObject('__ACTION');
	if(obj)
		obj.value = action;
};

claspCore.prototype.getObject = function (n, d) {
  var p,i,x;  

  if(!d) d=document; 
  	if((p=n.indexOf("?"))>0&&parent.frames.length) {  		
		d=parent.frames[n.substring(p+1)].document; 
		n=n.substring(0,p);		
	}
	if(!(x=d[n])&&d.all) {
		x=d.all[n]; 
		
	}
	for (i=0;!x&&i<d.forms.length;i++) {
		x=d.forms[i][n];		
	}
	for(i=0;!x&&d.layers&&i<d.layers.length;i++) {
		x=clasp.getObject(n,d.layers[i].document); 
		
	}	
	for(i=0;!x&&d.layers&&i<document.links.length;i++)		
		if(document.links[i].href.indexOf("'" + n + "'")>0) {
			x = document.links[i];
			break;
		}	
	//alert(n + ":" + x)	
	if(!x) { //alert(n)
		if(clasp.events.objcache)
			x = eval("clasp.events.objcache." + n);
	}
	
	if(!x) x=eval("document.all." + n);
						
	return x;
}

claspCore.prototype.openInFrame = function(fn,url){
	var myname;
	var fr = this.getFrame(fn);
	if (url && fr) fr.location = url;
	else {
		// load me inside the correct frame
		url = self.location;
		myname = self.name;
		if (fr && fn!=myname){
			fr.location = url;
			self.location = "about:blank";
		}
	}
};

claspCore.prototype.getFrame = 	function (fn){
	var f=document.frames[fn];
	if (!f) f= parent.frames[fn];
	if (!f) f= window.frames[fn];
	if (!f && window.frames['top']) f= window.frames['top'].frames[fn];
	return f;
};


claspForm = function() {
			this.loaded=true;			
			this.version = "1.0";
			this.name  = "CLASPForm"; //this is the name I want to use...
			this.name2 = "frmForm";   //support current versions
			this.validator = null;
}

claspForm.prototype.doPostBack = doPostBack;

claspForm.prototype.resetScroll = function(noRetry) {	
	if(!document.layers)  {		
	  var frm = clasp.form.getForm()
	  document.body.scrollTop  = frm.__SCROLLTOP.value;
	  document.body.scrollLeft = frm.__SCROLLLEFT.value;
	}
	
	// retry 10ms after page if scroll didn't reset
	if (!noRetry && (document.body.scrollTop!=frm.__SCROLLTOP.value)) {
		window.setTimeout('clasp.form.resetScroll(true)',10);
	}
}

claspForm.prototype.getForm = function() {	
	if(!this.form) {
		this.form = eval("document." + this.name);
		if(!this.form)
			this.form = eval("document." + this.name2);
	}
	return this.form;
}

claspForm.prototype.getField = function(fieldName) {
	var frm = this.getForm();
	var fld = eval("frm." + fieldName);
	return fld;
}

claspForm.prototype.enableField = function(fieldName) {
	var fld = this.getField(fieldName);
	fld.disabled = false;
}

claspForm.prototype.disableField = function(fieldName) {
	var fld = this.getField(fieldName);
	fld.disabled = true;
}

claspForm.prototype.isEnabled = function(fieldName) {
	var fld = this.getField(fieldName);	
	if(fld.disabled == null)
		return true;
	else
		return !fld.disabled;
}

claspForm.prototype.isDate = function(fieldName) {
	var fld  = this.getField(fieldName);
	var tst  = new Date(fld.value||'');
	return(!isNaN(tst));	
}

claspForm.prototype.isNumeric = function(fieldName) {
	var fld  = this.getField(fieldName);
	var tst  = new Number(fld.value||'');
	return(!isNaN(tst));
}

claspForm.prototype.hasValue = function(fieldName,def) {
	var fld = this.getField(fieldName);
	if(fld) {
		switch(fld.type) {
			case "text":			return (fld.value!="" && fld.value!=def);
			case "textarea":		return (fld.value!="" && fld.value!=def);
			case "password":		return (fld.value!="" && fld.value!=def);
			case "checkbox":		return fld.checked;
			case "select-one":		return (fld.value!="" && fld.value!=def)
			case "select-multiple":	return fld.selectedIndex>0;
			case "radio":			return fld.checked;
			case "hidden":			return (fld.value!="" && fld.value!=def);
			default:				for(var i=0;i<fld.length;i++) if(fld[i].checked)return true; return false;
		}			
	}
	return false;
}

claspForm.prototype.getValue = function(fieldName,def) {
	var fld = this.getField(fieldName);
	if(def==null)
		def = "";
	if(fld) {
		switch(fld.type) {
			case "text":			return fld.value?fld.value:def;
			case "textarea":		return fld.value?fld.value:def;
			case "password":		return fld.value?fld.value:def;
			case "checkbox":		return fld.checked;
			case "select-one":		return fld.value?fld.value:def;
			case "select-multiple":	return fld[fld.selectedIndex].value;//return fld.selectedIndex;
			case "radio":			return fld.checked;
			case "hidden":          return fld.value;
			default:			
					var	result = "";
					for(var i=0;i<fld.length;i++) 
							if(fld[i].checked)
								result = result + (result?',':'') + escape(fld[i].value);
					return result;
		}					
	}
	return null;
}

claspForm.prototype.isFieldPresent= function(fieldName) {
	var fld = this.getField(fieldName);
	return fld!=null;
}

claspForm.prototype.isClaspField= function(fieldName) {	
	return ( fieldName == "__FVERSION" || fieldName == "__ACTION"   || fieldName == "__SOURCE" ||
			 fieldName == "__INSTANCE" || fieldName == "__EXTRAMSG" || fieldName == "__ISPOSTBACK" ||
			 fieldName == "__ISREDIRECTEDPOSTBACK" || fieldName == "__SCROLLTOP" ||  fieldName == "__SCROLLLEFT" ||
			 fieldName == "__VIEWSTATE" || fieldName == "__VIEWSTATECOMPRESSED" )
}

claspForm.prototype.disableFieldsIn = function(containerID) {
	var container = document.all[containerID];
	if(container) {
		if(container)container.disabled = true;
		for(var x=0;x<container.all.length;x++)
			if(container.all[x].type) 
				container.all[x].disabled = true;		
	}
}
claspForm.prototype.enableFieldsIn = function(containerID) {
	var container = document.all[containerID];
	if(container) {
		if(container)container.disabled = false;
		for(var x=0;x<container.all.length;x++)
			if(container.all[x].type) 
				container.all[x].disabled = false;		
	}
}

claspForm.prototype.getXml = function() {
		var frm = clasp.form.getForm();
		var fieldName = '';
		var xml = '<fields>\n';
		var containerForm = new Object();
		var mx = frm.length;
		for(x=0;x<mx;x++) {
			fieldName = frm[x].name;
			var value = clasp.form.getValue(fieldName);
			
			if(value!=null && !eval('containerForm.' + fieldName)) {
				eval('containerForm.' + fieldName + '=1')
				xml = xml + "<" + fieldName + ">" ;
				if(fieldName=='__VIEWSTATE')
					xml = xml + value;
				else 
					xml = xml + escape(value);
				xml = xml + '</' + fieldName + '>\n';
			}							
		}
		xml = xml + '</fields>';
		containerForm = null;
		return xml;
}


//BACKWARD FUNCTIONALITY
	function doSubmit() {		
		alert('should not execute!');
		return false
	}	

	function doPostBack(action,object,instance,xmsg,frmaction,frameName) {
				
		var frm = document.frmForm;
		if(frm.__onSubmit) 
			if(!frm.__onSubmit()) return;								
			
		if(clasp.form.validator) {			
			var o = clasp.getObject(object);
			if(o) 
				if(!o.skipvalidation) 
				   if(!clasp.form.validator(o.validationgroup)) 
					  return;			
		}
		
		if(typeof(instance)=='object') {				
				if(instance.type == 'select-one' || instance.type == 'select-multiple')
					frm.__INSTANCE.value   = instance.selectedIndex;	
		}
		else {
			frm.__INSTANCE.value   = instance?instance:'0';	
		}
	
		frm.__ISPOSTBACK.value = "True";
		frm.__ACTION.value     = action;
		frm.__SOURCE.value     = object;
		frm.__EXTRAMSG.value   = xmsg;
		
		//To Restore Scroll Position
		if(!document.layers) {
			frm.__SCROLLTOP.value  = document.body.scrollTop;
			frm.__SCROLLLEFT.value = document.body.scrollLeft;
		}
		else {
			frm.__SCROLLTOP.value  = 0;
			frm.__SCROLLLEFT.value = 0;				
		}
		
		//Netscape 4.7 hmmm
		if(document.layers && !frmaction) {
			frm.action  = window.location;		
		}
		
		if(frmaction) { 
			frm.action = frmaction;
			frm.__ISREDIRECTEDPOSTBACK.value = 1
		}
		
		//Opens in another frame?
		if(frameName)
			frm.target = frameName;
		frm.submit();	
		return;
	}
	
//
// BrowserCheck
function BrowserCheck() {
	this.ns4 = (document.layers)? true:false;
	this.ie = (document.all&&(!window.opera))? true:false;
	this.dom = (document.getElementById)? true:false;
	this.ns6 = (window.sidebar)? true:false;
	this.moz = (window.sidebar||navigator.userAgent.indexOf('Gecko')!=-1)? true:false;
	this.opera = (window.opera)? true:false;
	this.mac = (navigator.userAgent.indexOf('Mac')!=-1)? true:false;
}
	
// Page Object	
claspPage = function() {
	this.loaded = true;
	this.version = "1.0";
	this.WebControls = new Object();
	this.WebControls.all   = [];
	this.WebControls._idcnt = 0;
	this.AjaxProcessing = false;
};

claspPage.prototype.addControl = function(name) {
	eval("clasp.page." + name + " = new WebControl('" + name + "')");
};

// Ajax Implementation
claspAjax = function() {
	this.version = "1.0";
	this.loaded  = true;
	this.requests = [];
	this.debug    = true;
}

claspAjax.prototype.getXMLHttpRequest = function() {
	if(clasp.ie)
		return new ActiveXObject("Microsoft.XMLHTTP");
	if(typeof(XMLHttpRequest)) 
		return new XMLHttpRequest;
	return null;
}

claspAjax.prototype.invoke = function(action,async,callbackFnc,actionSrc,executeOnly,destUrl) {	
	var x = 0;
	for(x=0;x<this.requests.length;x++) {
		if(this.requests[x].Available)
			return this.requests[x].invoke(action,async,callbackFnc,actionSrc,executeOnly,destUrl);
	}
	var obj = new claspAjaxCall(this.requests.length)
	this.requests.push(obj);
	return obj.invoke(action,async,callbackFnc,actionSrc,executeOnly,destUrl);
}

claspAjaxCall = function(index) {
	this.Available   = true;
	this.httpRequest = null;
	this.callbackFnc = null;
	this.index       = index;
}

claspAjaxCall.prototype.onCallBack = function(data) {
	this.Available   = true;
	if(clasp.ajax.debug) window.status = 'request (' + this.index + ',' + this.action + ') completed at ' + (new Date());
	if(this.callbackFnc)
		return this.callbackFnc(data);
	else
		return data;
};

claspAjaxCall.prototype.onReadyStateChange = function(async)  {
	if(clasp.ajax.debug) window.status = 'processing request (' + this.index + ') ' + (new Date());
	try
	{	if(this.httpRequest.readyState==4){
			var ret = this.httpRequest.responseText;
			if(ret.toString().indexOf('ERR$')>=0)
			{	var msg = ret.substring(4, ret.length)
				var err =  new Error()
				err.message = 'Ajax Server Error : ' + msg
				throw err;
			}else
				if(async) {
					var resp = this.httpRequest.responseText.split("|*|");
					if(resp[1])	clasp.getObject("__VIEWSTATE").value  = resp[1];			
					return this.onCallBack(resp[0]);				
				}
		}
	}
	catch(ex)
	{   alert('Error OnReadyStateChange: ' + ex.message)
		return false;
	}
};

claspAjaxCall.prototype.invoke = function(action,async,callbackFnc,actionSrc,executeOnly,processorUrl) {	
	var callBackIndex = this.index;
	var destUrl = processorUrl?processorUrl:window.location;	
	this.action = action;
	try {
		if(!this.httpRequest) {			
			this.httpRequest= clasp.ajax.getXMLHttpRequest();
		}
		if(!this.Available)	{
			alert('Request in progress, please wait...');
			return false;
		}			
		
		this.Available = false;

		this.callbackFnc = callbackFnc;
		this.httpRequest.open('POST', destUrl, async?true:false);
		this.httpRequest.setRequestHeader('Content-Type','application/x-www-form-urlencoded');

		this.httpRequest.onreadystatechange = function () { 				
			clasp.ajax.requests[callBackIndex].onReadyStateChange(async);
		}	

		var request = new Array();
		var req = 'AjaxPost=1&AjaxAction=' + action + '&AjaxSource=' + actionSrc + '&AjaxMode=' + (executeOnly?1:0)

		if(!executeOnly) 
			req = req + "&AjaxForm=" + clasp.form.getXml();
		this.httpRequest.send(req);			
		
		if(!async) {
			var resp = this.httpRequest.responseText.split("|*|");
			if(resp[1])	clasp.getObject("__VIEWSTATE").value  = resp[1];			
			return this.onCallBack(resp[0]);
		}		
	}catch(ex) {
		var ret = this.httpRequest.responseText;
		alert('Invoke: ' +  ex.message + ':' + ret);
	}
};//

// WebControls client side
WebControl = function(name) {
	this.ctrlName = name;
	this.id = name;
};
WebControl.prototype.invokeEvent = function(action,instance,xmsg,frmaction,frameName) {
	clasp.form.doPostBack(action,this.ctrlName,instance,xmsg,frmaction,frameName);
};

WebControl.prototype.toString = function(){
	return "clasp.page." + this.id;
};

WebControl.prototype.obj = function() {
	return clasp.getObject(this.id);
}

WebControl.prototype.onCallBack = function(data) {	
	if(this.callBack)
		return this.callBack(data)
	else 
		return data;
}

WebControl.prototype.invoke = function(params) {
	return clasp.ajax.invoke(this.id + '_OnClientEvent',false,this.onCallBack,this.id,false);	
}

//Initialize Client Side JS
clasp = new claspCore()
clasp.ua = new BrowserCheck();
clasp.loadLibrary("events");
