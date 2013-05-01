claspdhtml = function() {
	this.loaded  = true;
	this.version = "1.0";
}
clasp.dhtml = new claspdhtml();


claspdhtml.prototype.left  = function (elmId) {
	xPos = eval(elmId).offsetLeft;
	var tempEl = eval(elmId).offsetParent;
  	while (tempEl != null) {
  		xPos += tempEl.offsetLeft;
  		tempEl = tempEl.offsetParent;
  	}
	return xPos;
}

claspdhtml.prototype.top = function (elmId) {
	yPos = eval(elmId).offsetTop;
	tempEl = eval(elmId).offsetParent;
	while (tempEl != null) {
  		yPos += tempEl.offsetTop;
  		tempEl = tempEl.offsetParent;
  	}
	return yPos;
}


claspdhtml.prototype.moveY = function(id, increment) {
  if (clasp.ns4) eval(id).top += increment
  else eval(id).style.pixelTop += increment;
}

claspdhtml.prototype.getTop = function(id) {
  if (clasp.ns4) return eval(id).top
  else return eval(id).style.pixelTop;
}

claspdhtml.prototype.setTop = function(id, top) {
  if (clasp.ns4) eval(id).top   = top
  else eval(id).style.top = top;
}

claspdhtml.prototype.setLeft = function(id, left) {
  if (clasp.ns4) eval(id).left   = left
  else eval(id).style.left = left;
}

claspdhtml.prototype.bgColor = function(id, color) {
  if (clasp.ns4) eval(id).bgColor = color
  else eval(id).style.backgroundColor = color;
}

claspdhtml.prototype.createClipRegionAt00 = function(id, clipWidth, clipHeight) {
  if (clasp.ns4) {
    eval(id).clip.width = clipWidth;
    eval(id).clip.height = clipHeight;
  } else eval(id).style.clip = "rect(0 " + eval(clipWidth) + " " + eval(clipHeight) + " 0)";
}

claspdhtml.prototype.getHeight = function(id) {
  if (clasp.ns4) return eval(id).clip.height
  else return eval(id).clientHeight;
}

claspdhtml.prototype.getImageWidth = function(elmId) {
  return eval(elmId).width;
}

claspdhtml.prototype.getImageHeight = function(elmId) {
  return eval(elmId).height;
}

claspdhtml.prototype.getWindowWidth = function() {
  if (clasp.ns4) {return window.innerWidth}
  else {return document.body.clientWidth}
}

claspdhtml.prototype.getWindowHeight = function() {
  if (clasp.ns4) {return window.innerHeight}
  else {return document.body.clientHeight}
}

claspdhtml.prototype.getPageScrollLeft = function() {
  if (clasp.ns4) {return window.pageXOffset}
  else {return document.body.scrollLeft}
}

claspdhtml.prototype.getPageScrollTop = function() {
  if (clasp.ns4) {return window.pageYOffset}
  else {return document.body.scrollTop}
}


claspdhtml.prototype.getWidth = function(id) {
  if (clasp.ns4) {return eval("document." + id + ".clip.width")}
  else {return eval("document.all." + id + ".style.pixelWidth")}
}

claspdhtml.prototype.getVisibility = function(id) {
  if (clasp.ns4) {
    if (eval("document." + id).visibility == "show") return true
    else return false;
  }
  else {
   if (eval("document.all." + id).style.visibility == "visible") return true
    else return false;
  }
}

claspdhtml.prototype.setVisibility = function(id, flag) {
  if (clasp.ns4) {
    var str = (flag) ? 'show' : 'hide';
    eval("document." + id).visibility = str;
  }
  else {
    var str = (flag) ? 'visible' : 'hidden';
    eval("document.all." + id).style.visibility = str;
  }
}

claspdhtml.prototype.setElementVisibility = function(id, flag) {
  if (clasp.ns4) {
    var str = (flag) ? 'show' : 'hide';
    eval(id).visibility = str;
  }
  else {
    var str = (flag) ? 'visible' : 'hidden';
    eval(id).style.visibility = str;
  }
}

claspdhtml.prototype.setPosFromLeft = function(id, xCoord) {
  if (clasp.ns4) {eval("document." + id).left = xCoord}
  else {eval("document.all." + id).style.left = xCoord}
}

claspdhtml.prototype.setElementLeft = function(id, elementLeft) {
  if (clasp.ns4) eval(id).left = elementLeft
  else eval(id).style.left = elementLeft;
}

claspdhtml.prototype.setElementWidth = function(id, elementWidth) {
  if (clasp.ns4) eval(id).width = elementWidth
  else eval(id).style.width = elementWidth;
}

claspdhtml.prototype.setElementHeight = function(id, elementHeight) {
  if (clasp.ns4) eval(id).height = elementHeight
  else eval(id).style.height = elementHeight;
}

claspdhtml.prototype.setPosFromTop = function(id, yCoord) {
  if (clasp.ns4) {eval("document." + id).top = yCoord}
  else {eval("document.all." + id).style.top = yCoord}
}

claspdhtml.prototype.setElementTop = function(id, elementTop) {
  if (clasp.ns4) eval(id).top = elementTop
  else eval(id).style.top = elementTop;
}

claspdhtml.prototype.setZposition = function(id, z) {
  if (clasp.ns4) {eval("document." + id).zIndex = z}
  else {eval("document.all." + id).style.zIndex = z}
}

claspdhtml.prototype.findHighestZ = function() {
  var documentDivs = new Array();
  if (clasp.ns4) {documentDivs = document.layers}
  else {documentDivs = document.all.tags("DIV")};
  var highestZ = 0;
  for (var i = 0; i < documentDivs.length; i++) {
     var zIndex = (clasp.ns4) ? documentDivs[i].zIndex : documentDivs[i].style.zIndex;
     highestZ = (zIndex > highestZ) ? zIndex : highestZ;
  }
  return highestZ;
}

