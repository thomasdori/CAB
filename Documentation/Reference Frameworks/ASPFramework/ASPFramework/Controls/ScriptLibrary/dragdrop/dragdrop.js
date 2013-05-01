/**
 * Drag & Drop Library
 * 
*/

var bV=parseInt(navigator.appVersion);
NS4=(document.layers) ? true : false;
IE4=((document.all)&&(bV>=4))?true:false;

currentX = currentY = 0;
DragElm = null;
ActiveElm = null;


function initNS4(args){

	// load arguments	
	if (!args) args={};
	var e = args.e;
	var elm = args.elm;
	var ctrl = args.ctrl;
	var pb = args.pb;	// prevent bubble

	if (!elm) return null;
	elm.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP | Event.MOUSEMOVE)
	elm.onmouseup = function dropNS4(e) {
		return dropElm({'elm':this,'ctrl':ctrl,'pb':pb});
	};
	elm.onmousedown = function(e) {
		return grabElm({'e':e,'elm':this,'pb':pb});
	};
	elm.onmousemove = function(e) {
		return moveElm({'e':e,'elm':this,'pb':pb});
	};
};

function grabElm(args) {

	// load arguments	
	if (!args) args={};
	var e = args.e;
	var elm = args.elm;
	var pb = args.pb;	// prevent bubble

    if (IE4) {
        if(elm) DragElm = elm;
        else {
        	DragElm = event.srcElement;
			while (DragElm.id.indexOf("DRAG") == -1) {
				DragElm = DragElm.parentElement;
				if (DragElm == null) { return }
			}
		}
    }
    else {
        mouseX = e.pageX;
        mouseY = e.pageY;
        
        if(elm) DragElm = elm;
        else {
			for ( i=0; i<document.layers.length; i++ ) {
				tempLayer = document.layers[i];
				if ( tempLayer.id.indexOf("DRAG") == -1 ) { continue }
				if ( (mouseX > tempLayer.left) && (mouseX < (tempLayer.left + tempLayer.clip.width)) && (mouseY > tempLayer.top) && (mouseY < (tempLayer.top + tempLayer.clip.height)) ) {
					DragElm = tempLayer;
				}
			} 
		}
        if (DragElm == null) { return}
    }
	
	if (ActiveElm==null) ActiveElm = DragElm; // ?
	if (DragElm != ActiveElm) {
		if (IE4) { DragElm.style.zIndex = ActiveElm.style.zIndex+1 }
		else { DragElm.moveAbove(ActiveElm) };
		ActiveElm = DragElm;
	}
	
    if (IE4) {
        DragElm.style.pixelLeft = DragElm.offsetLeft;
        DragElm.style.pixelTop = DragElm.offsetTop;

        currentX = (e.clientX + document.body.scrollLeft);
        currentY = (e.clientY + document.body.scrollTop); 

    }
    else {
		currentX = e.pageX;
		currentY = e.pageY;
		document.captureEvents(Event.MOUSEMOVE);
		document.onmousemove = moveElm;
    }
    if (pb) return pbElm(e);    
};


function moveElm(args) {

	// load arguments	
	if (!args) args={};
	var e = args.e;
	var pb = args.pb;	// prevent bubble

    if (DragElm == null) { return };

    if (IE4) {
        newX = (event.clientX + document.body.scrollLeft);
        newY = (event.clientY + document.body.scrollTop);
    }
    else {
        newX = e.pageX;
        newY = e.pageY;
    }
    distanceX = (newX - currentX);
    distanceY = (newY - currentY);
    currentX = newX;
    currentY = newY;

    if (IE4) {
        DragElm.style.pixelLeft += distanceX;
        DragElm.style.pixelTop += distanceY;
        e.returnValue = false;
    }
    else { 
    	DragElm.moveBy(distanceX,distanceY) 
    }    
    if (pb)  return pbElm(e);
};

function dropElm(args) {
	var css;
	var l,t,z;

	// load arguments	
	if (!args) args={};
	var e = args.e;
	var elm = args.elm;
	var ctrl = args.ctrl;
	var pb = args.pb;	// prevent bubble
	
    if (NS4 && elm) { elm.releaseEvents(Event.MOUSEMOVE) }

	if (ctrl && DragElm) {
		if(NS4) css = DragElm
		else css = DragElm.style;

		l = parseInt(css.left);
		t = parseInt(css.top);
		z = css.zIndex;

		// update control's left, top and zindex position
		var c = (ctrl)? clasp.form.getField(ctrl):null;
		if (c) {
			c.value = l+";"+t+";"+z; // save position
		}
	}
    DragElm = null;
    if (pb)  return pbElm(e);
};

function checkElm() {
    if (DragElm!=null) { return false }
};

// Prevent Bubble
function pbElm(e){
	e.cancelBubble = true;
	return false;
}