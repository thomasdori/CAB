//

function CLASP_ProgressBar(Id,min,max,value,width,height) {
	this.Id    = Id;
	this.min   = min;
	this.max   = max;
	this.value = value
	this.loaded = false;
	this.width   = width?width:200;
	this.height  = height?height:20;
	this.update = CLAS_ProgressBar_Update;
	this.reset  = CLAS_ProgressBar_Reset;
	this.init   = CLAS_ProgressBar_Init;
}

function CLAS_ProgressBar_Init() {
	this.span1 = document.getElementById(this.Id)
	this.span2 = document.getElementById(this.Id + "_2")
	if(this.span2) {
		this.width = this.span2.style.pixelWidth;	
	}
	this.steps = this.width / this.max;
	this.loaded = true;
	if(this.Id.style) this.span1=this.Id;
}

function CLAS_ProgressBar_Update(value) {
	
	if(!this.loaded) this.init();	
	if(value) {
		this.value = value;
	}	
	this.span1.style.pixelWidth += (this.value * this.steps)
}
function CLAS_ProgressBar_Reset() {
}
