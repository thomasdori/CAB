//BY CJCA. Used to attach and bubble events... nice!.
claspEvent = function (obj,eventName,newEvent,isCode,which) {
	this.active = true;
	this.index = clasp.events.handlers.length;	
	this.obj = obj;		
	if(!this.oldhandler) {		
		if(!this.obj.type && this.obj.href && this.obj.href != "#" && eventName.toLowerCase() == "onclick") {
			this.oldhandler  = new Function(this.obj.href.replace("javascript:",""));						
			this.obj.href = "#";
		}else {			
			if( eval("this.obj." + eventName ) )
				this.oldhandler = eval("this.obj." + eventName);
		}
	}else
		this.oldhandler = eval("this.obj." + eventName);		
	
	this.handler = this.getFunctionPointer(newEvent);
	eval("this.obj." + eventName + " = new Function(\" clasp.events.handlers[" + this.index + "].handle(); \")");		
};

claspEvent.prototype.emptyFunction = function() {
	alert("uh?");
}

claspEvent.prototype.getFunctionPointer = function(e) {
	var evt;	
	if(typeof(e) == "function")
		return e;	
	if(typeof(e) == "string") {	
		if(e.indexOf("'")>0 || e.indexOf("(")>0 || e.indexOf(";")>0 || e.indexOf("=")>0) {
			return eval( " new Function('" + e.replace(/\'/ig,"\\'") + "')" );			
		}
		else
			return eval(e);
	}
}

claspEvent.prototype.createEvent = function(eventCode) {
	return eval( " new Function() { " + eventCode + " } " )	
}

claspEvent.prototype.handle = function() {			
	if (this.handler())  {		
		if(this.oldhandler) this.oldhandler();		
	}
}

claspEvents = function() {
	this.loaded = true;
	this.version = "1.0";		
	this.handlers = [];
	this.queued_handlers = []
	this.queuing = false;
	this.queue_processed = false;
}

claspEvents.prototype.addListener = function(object,eventName,newEvent) {		

	if(this.queue_processed)
		return;	

	if(!this.queuing) {
		this.queuing = true;				
		if(window.attachEvent)
			window.attachEvent("onload",clasp.events.processQueue);
		else {
			window.captureEvents("Load");
			window.onload = clasp.events.processQueue;
		}
	}	
	this.queued_handlers.push ( new Array(object,eventName,newEvent)  ) 
}

claspEvents.prototype.processQueue = function() {
	var onloadpresent = false;
	var onload_index;
	if(window.detachEvent)
		window.detachEvent("onload",clasp.events.processQueue);
	else 
		window.releaseEvents("Load");	
	clasp.events.queue_processed = true;
	clasp.events.queuing = false;
	clasp.events.objcache = new Object(); //important for ns4 & link/image buttons...
	
	for(x=0;x<clasp.events.queued_handlers.length;x++) {
		var e = clasp.events.queued_handlers[x];		
		var obj;
		if(e[1] != 'onload') {	
			if(typeof(e[0])=="string")  { 
				obj = eval("clasp.events.objcache." + e[0]);
				if(!obj) obj = clasp.getObject(e[0])
				eval("clasp.events.objcache." + e[0] + " = obj ")	
			}else 	
				obj = e[0];						   			
			clasp.events.handlers.push( new claspEvent(obj,e[1],e[2]) );
		}else {
			onloadpresent = true;
			clasp.events.handlers.push(new claspEvent(document.body,"onload",e[2]))
			onload_index = clasp.events.handlers.length - 1;
		}			
	}
	if(onloadpresent) {
		setTimeout("clasp.events.handlers[" + onload_index + "].handle()",100)
		//setTimeout("document.body.onload()",100); //this should work... is not however! :-(
	}
}

claspEvents.prototype.removeListener = function (objectID,eventName,newEvent) {
	//hmmm... I may have to do something here.
	//get new event...
}

claspEvents.prototype.setFocus = function(obj) {
	if(arguments[0]) {
		var o = clasp.getObject(obj);
		/*if(o) 
			if(o.length && o.length>=arguments[1]) 					
				eval("o[" + arguments[1] + "].focus()");
			else
				o.focus();
		else*/
			if(o.focus) o.focus();
	}
	return true;
}

clasp.events = new claspEvents;