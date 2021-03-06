
/*===================================================================
 Author: Matt Kruse
 
 View documentation, examples, and source code at:
     http://www.JavascriptToolbox.com/

 NOTICE: You may use this code for any purpose, commercial or
 private, without any further permission from the author. You may
 remove this notice from your final code if you wish, however it is
 appreciated by the author if at least the web site address is kept.

 This code may NOT be distributed for download from script sites, 
 open source CDs or sites, or any other distribution method. If you
 wish you share this code with others, please direct them to the 
 web site above.
 
 Pleae do not link directly to the .js files on the server above. Copy
 the files to your own server for use with your site or webapp.
 ===================================================================*/
/*
This code is inspired by and extended from Stuart Langridge's aqlist code:
		http://www.kryogenix.org/code/browser/aqlists/
		Stuart Langridge, November 2002
		sil@kryogenix.org
		Inspired by Aaron's labels.js (http://youngpup.net/demos/labels/) 
		and Dave Lindquist's menuDropDown.js (http://www.gazingus.org/dhtml/?id=109)
*/
var converted=false;
// Automatically attach a listener to the window onload, to convert the trees
// blocked here so we can call it inline - then open the nodes we need
//addEvent(window,"load",convertTrees);

// Utility function to add an event listener
function addEvent(o,e,f){
	if (o.addEventListener){ o.addEventListener(e,f,false); return true; }
	else if (o.attachEvent){ return o.attachEvent("on"+e,f); }
	else { return false; }
}

// utility function to set a global variable if it is not already set
function setDefault(name,val) {
	if (typeof(window[name])=="undefined" || window[name]==null) {
		window[name]=val;
	}
}

// Full expands a tree with a given ID
function expandTree(treeId) {
	var ul = document.getElementById(treeId);
	if (ul == null) { return false; }
	expandCollapseList(ul,nodeOpenClass);
}

// Fully collapses a tree with a given ID
function collapseTree(treeId) {
	var ul = document.getElementById(treeId);
	if (ul == null) { return false; }
	expandCollapseList(ul,nodeClosedClass);
}

// Expands enough nodes to expose an LI with a given ID
function expandToItem(treeId,itemId) {
	//alert('expandToItem, treeId='+treeId+', itemId='+itemId);
	var ul = document.getElementById(treeId);
	if (ul == null) { 
		//alert('ul not found');
		return false; 
	}
	var ret = expandCollapseList(ul,nodeOpenClass,itemId);
	if (ret) {
		var o = document.getElementById(itemId);
		if (o.scrollIntoView) {
			o.scrollIntoView(false);
		}
	}
}

// Performs 3 functions:
// a) Expand all nodes
// b) Collapse all nodes
// c) Expand all nodes to reach a certain ID
function expandCollapseList(ul,cName,itemId) {
	//alert('expandCollapseList, cName='+cName+', itemId='+itemId);
	if (!ul.childNodes || ul.childNodes.length==0) { return false; }
	// Iterate LIs
	for (var itemi=0;itemi<ul.childNodes.length;itemi++) {
		var item = ul.childNodes[itemi];
		if (itemId!=null && item.id==itemId) {
			item.className=cName;
                        return true; 
                        }
		//alert('item.nodeName='+item.nodeName);
		if (item.nodeName == "LI") {
			// Iterate things in this LI
			var subLists = false;
			for (var sitemi=0;sitemi<item.childNodes.length;sitemi++) {
				var sitem = item.childNodes[sitemi];
				if (sitem.nodeName=="UL") {
					subLists = true;
					var ret = expandCollapseList(sitem,cName,itemId);
					if (itemId!=null && ret) {
						item.className=cName;
						return true;
					}
				}
			}
			if (subLists && itemId==null) {
				item.className = cName;
			}
		}
	}
}

// Search the document for UL elements with the correct CLASS name, then process them
function convertTrees() {
	setDefault("treeClass","mktree");
	setDefault("nodeClosedClass","mklc");
	setDefault("nodeOpenClass","mklo");
	setDefault("nodeBulletClass","mklb");
	setDefault("nodeLinkClass","mkb");
	setDefault("preProcessTrees",true);
	return true;
	if (preProcessTrees) {
		if (!document.createElement) { return; } // Without createElement, we can't do anything
		var uls = document.getElementsByTagName("ul");
		if (uls==null) { return; }
		var uls_length = uls.length;
		for (var uli=0;uli<uls_length;uli++) {
			var ul=uls[uli];
			if (ul.nodeName=="UL" && ul.className==treeClass) {
				processList(ul);
			}
		}
	}
}

function treeNodeOnclick() {
	this.parentNode.parentNode.parentNode.className = (this.parentNode.parentNode.parentNode.className==nodeOpenClass) ? nodeClosedClass : nodeOpenClass;
	return false;
}
function mkClick(e) {
	e.parentNode.parentNode.className = (e.parentNode.parentNode.className==nodeOpenClass) ? nodeClosedClass : nodeOpenClass;
	return false;
}

function retFalse() {
	return false;
}
// Process a UL tag and all its children, to convert to a tree
function processList(ul) {
	if (!ul.childNodes || ul.childNodes.length==0) { return; }
	// Iterate LIs
	var childNodesLength = ul.childNodes.length;
	for (var itemi=0;itemi<childNodesLength;itemi++) {
		var item = ul.childNodes[itemi];
		if (item.nodeName == "LI") {
			// Iterate things in this LI
			var subLists = false;
			var itemChildNodesLength = item.childNodes.length;
			for (var sitemi=0;sitemi<itemChildNodesLength;sitemi++) {
				var sitem = item.childNodes[sitemi];
				if (sitem.nodeName=="UL") {
					subLists = true;
					processList(sitem);
				}
			}
			var s= document.createElement("SPAN");
			var t= '\u00A0'; // &nbsp;
			s.className = nodeLinkClass;
			if (subLists) {
				// This LI has UL's in it, so it's a +/- node
				if (item.className==null || item.className=="") {
					item.className = nodeClosedClass;
				}
				// If it's just text, make the text work as the link also
				if (item.firstChild.nodeName=="#text") {
					t = t+item.firstChild.nodeValue;
					item.removeChild(item.firstChild);
				}
				s.onclick = treeNodeOnclick;
			}
			else {
				// No sublists, so it's just a bullet node
				item.className = nodeBulletClass;
				s.onclick = retFalse;
			}
			var da= document.createElement("DIV");
			var dr= document.createElement("DIV");
			da.style.position="absolute";
			dr.style.position="relative";
			dr.style.left = -20;
			s.appendChild(document.createTextNode(t));
			dr.appendChild(s);
			da.appendChild(dr);
			item.insertBefore(da,item.firstChild);
//			item.insertBefore(s,item.firstChild);
		}
	}
}
