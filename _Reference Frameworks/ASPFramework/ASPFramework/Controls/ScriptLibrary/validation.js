//Required Field
claspReqFieldValidator = function(name,def,msg,attachTo,group) {
	this.name = name;
	this.def  = def;
	this.msg  = msg;
	this.group = group;
	this.attachTo = attachTo;
	this.index = clasp.form.validation.add(this,group) - 1;	
	clasp.events.addListener(name,"onblur",new Function("clasp.form.validation.validateOne(" + this.index + (group ? ",'" + group + "'" :"") + ");return true;") )
}
claspReqFieldValidator.prototype.validate = function() {	
	if( clasp.form.hasValue(this.name,this.def) ) {
		if(this.attachTo)
			clasp.form.validation.setElemementVisibility(this.attachTo,false);
		return true;
	}	
	if(this.attachTo) {
		clasp.form.validation.setElemementVisibility(this.attachTo,true);
	}			
	return false;
}

claspDateFieldValidator = function(name,attachTo,group) {
	this.name = name;
	this.group = group;
	this.attachTo = attachTo;
	this.index = clasp.form.validation.add(this,group) - 1;	
	clasp.events.addListener(name,"onblur",new Function("clasp.form.validation.validateOne(" + this.index + (group ? ",'" + group + "'" :"") + ");return true;") )
}

claspDateFieldValidator.prototype.validate = function() {	
	if( clasp.form.hasValue(this.name,this.def) ) {
		var ok = clasp.form.isDate(this.name);
		if(!ok) {
			if(this.attachTo) clasp.form.validation.setElemementVisibility(this.attachTo,true);
			return false;
		}
	}
	if(this.attachTo)clasp.form.validation.setElemementVisibility(this.attachTo,false);
	return true;
}

claspNumericFieldValidator = function(name,attachTo,group) {
	this.name = name;
	this.group = group;
	this.attachTo = attachTo;
	this.index = clasp.form.validation.add(this,group) - 1;	
	clasp.events.addListener(name,"onblur",new Function("clasp.form.validation.validateOne(" + this.index + (group ? ",'" + group + "'" :"") + ");return true;") )
}
claspNumericFieldValidator.prototype.validate = function() {	
	if( clasp.form.hasValue(this.name,this.def) ) {
		var ok = clasp.form.isNumeric(this.name);
		if(!ok) {
			if(this.attachTo)clasp.form.validation.setElemementVisibility(this.attachTo,true);
			return false;
		}			
	}
	if(this.attachTo) clasp.form.validation.setElemementVisibility(this.attachTo,false);
	return true;
}



claspValidation = function() {
	this.loaded = true;
	this.version = "1.0";	
	this.groups  = new Array( {name:"default", fields:[],errorMsg:[]}  )
	clasp.form.validator = new Function('return clasp.form.validation.validateForm(arguments[0])');
}

claspValidation.prototype.setElemementVisibility = function(id, visible) {
  if (clasp.ns4)
    eval("document." + id).visibility = (visible) ? 'show' : 'hide';
  else
    eval("document.all." + id).style.visibility = (visible) ? 'visible' : 'hidden';
}

claspValidation.prototype.add = function (validator,groupName)  {
	var g;
	g = this.getgroup(groupName);	
	g.fields.push ( validator );
	return g.fields.length;
}

claspValidation.prototype.validateOne = function (i,gn)  {
	var g = gn?this.getgroup(gn):this.groups[0];
	var v = g.fields[i];
	if(v) v.validate();
}

claspValidation.prototype.validateForm = function()  {
	var gn = arguments?arguments[0]:""; //args[0] = groupname
	var g = clasp.form.validation.getgroup(gn);	
	
	var f;
	var r = true;
	for(x=0;x<g.fields.length;x++) {
		f = g.fields[x];
		r&= f.validate();
	}
	return r;
}

claspValidation.prototype.getgroup = function(name) {		
	if(!name || name == "")
		return this.groups[0];
	
	for(var x=0;x<this.groups.length;x++)
		if(this.groups[x].name == name)
			return this.groups[x];
				
	this.groups.push ( {name:name, fields:[]}  )	
	return this.groups[x];
}

clasp.form.validation	= new claspValidation();
