// Title: Tigra Tree
// Description: See the demo at url
// URL: http://www.softcomplex.com/products/tigra_menu_tree/
// Version: 1.1
// Date: 11-12-2002 (mm-dd-yyyy)
// Notes: This script is free. Visit official site for further details.


function tree (a_items, a_template,menu_name,items_open,selected_node) {

	this.a_tpl      = a_template;
	this.a_config   = a_items;
	this.o_root     = this;
	this.a_index    = [];
	this.o_selected = null;
	this.n_depth    = -1;
	//CLASP
		this.menu_name = menu_name;
		this.open_list = items_open.split(",");
		this.a_flat_nodes = [];
	//CLASP
		
	var o_icone = new Image(),
		o_iconl = new Image();
	o_icone.src = a_template['icon_e'];
	o_iconl.src = a_template['icon_l'];
	a_template['im_e'] = o_icone;
	a_template['im_l'] = o_iconl;
	for (var i = 0; i < 64; i++)
		if (a_template['icon_' + i]) {
			var o_icon = new Image();
			a_template['im_' + i] = o_icon;
			o_icon.src = a_template['icon_' + i];
		}
	
	this.toggle = function (n_id) {	var o_item = this.a_index[n_id]; o_item.open(o_item.b_opened) };
	this.select = function (n_id) { return this.a_index[n_id].select(); };
	this.mout   = function (n_id) { this.a_index[n_id].upstatus(true) };
	this.mover  = function (n_id) { this.a_index[n_id].upstatus() };
	
	this.a_children = [];
	for (var i = 0; i < a_items.length; i++)
		new tree_item(this, i);

	this.n_id = trees.length;
	trees[this.n_id] = this;
	
	for (var i = 0; i < this.a_children.length; i++) {
		document.write(this.a_children[i].init());
		this.a_children[i].open();
	}
		
	//CLASP
	if(this.open_list) {
		tree_restore_opened_items(this);
	}
	if(selected_node) {
		this.a_flat_nodes[selected_node].ref.select(false,true);
	}	
	if(this.n_id == 0)
		clasp.events.addListener(clasp.form.getForm(),"submit",tree_persis_treeview_status)
	//CLASP

}
function tree_item (o_parent, n_order) {

	this.n_depth  = o_parent.n_depth + 1;
	this.a_config = o_parent.a_config[n_order + (this.n_depth ? 2 : 0)];
	if (!this.a_config) return;

	this.o_root    = o_parent.o_root;
	this.o_parent  = o_parent;
	this.n_order   = n_order;
	this.b_opened  = !this.n_depth;

	this.n_id = this.o_root.a_index.length;
	this.o_root.a_index[this.n_id] = this;
	o_parent.a_children[n_order] = this;
	
	this.o_root.a_flat_nodes.push({ref:this});//CLASP

	this.a_children = [];
	for (var i = 0; i < this.a_config.length - 2; i++)
		new tree_item(this, i);

	this.get_icon = item_get_icon;
	this.open     = item_open;
	this.select   = item_select;
	this.init     = item_init;
	this.upstatus = item_upstatus;
	this.is_last  = function () { return this.n_order == this.o_parent.a_children.length - 1 };
	this.initialized = false;
}

function item_open (b_close) {

	var o_idiv = get_element('i_div' + this.o_root.n_id + '_' + this.n_id);
	if (!o_idiv) return;
	
	if (!o_idiv.innerHTML) {
		var a_children = [];
		for (var i = 0; i < this.a_children.length; i++)
			a_children[i]= this.a_children[i].init();
		o_idiv.innerHTML = a_children.join('');
	}
	o_idiv.style.display = (b_close ? 'none' : 'block');
	this.b_opened = !b_close;
	var o_jicon = document.images['j_img' + this.o_root.n_id + '_' + this.n_id],
		o_iicon = document.images['i_img' + this.o_root.n_id + '_' + this.n_id];
	if (o_jicon) o_jicon.src = this.get_icon(true);
	if (o_iicon) o_iicon.src = this.get_icon();
	this.upstatus();
}

function item_select (b_deselect,b_bypass) {
	if (!b_deselect) {
		var o_olditem = this.o_root.o_selected;
		this.o_root.o_selected = this;
		if (o_olditem) o_olditem.select(true);		
		//CLASP INTEGRATION		
		var item_action = this.a_config[1];
		var menu_name   = this.o_root.menu_name;
		if(!item_action) item_action = "";
		//If no action or action contains clasp and is a leaf
		if( (item_action.indexOf("clasp:")>=0  || (item_action == "" && this.a_children.length==0) ) && !b_bypass)  {
			var action = "" + item_action.replace("clasp:","");
			var evt    = "OnClick";
			if(action!="") 	evt = action;			
			clasp.form.doPostBack(evt,menu_name,this.n_id);			
			return false;
		}//END CLASP INTEGRATION
	}
	
	var o_iicon = document.images['i_img' + this.o_root.n_id + '_' + this.n_id];
	if (o_iicon) o_iicon.src = this.get_icon();
	//CLASP
	var elm = get_element('i_txt' + this.o_root.n_id + '_' + this.n_id);
	if(elm)
		elm.style.fontWeight = b_deselect ? 'normal' : 'bold';
	//CLASP
	this.upstatus();
	if(!b_deselect && this.a_config[1]!=null)
	//if(this.a_config[1].indexOf("js:")>=0) {
	if(this.a_config[1].indexOf("javascript:")>=0) {
		//var action = "" + this.a_config[1].replace("js:","");
		//alert(this.a_config[1])
		return eval(this.a_config[1]);
	}

	return Boolean(this.a_config[1]);
}

function item_upstatus (b_clear) {
	window.setTimeout('window.status="' + (b_clear ? '' : this.a_config[0] + (this.a_config[1] ? ' ('+ this.a_config[1] + ')' : '')) + '"', 10);
}

function item_init () {
	
	var a_offset = [],
		o_current_item = this.o_parent;
	for (var i = this.n_depth; i > 1; i--) {
		a_offset[i] = '<img src="' + this.o_root.a_tpl[o_current_item.is_last() ? 'icon_e' : 'icon_l'] + '" border="0" align="absbottom">';
		o_current_item = o_current_item.o_parent;
	}
	return '<table cellpadding="0" cellspacing="0" border="0"><tr><td nowrap>' + (this.n_depth ? a_offset.join('') + (this.a_children.length
		? '<a href="javascript: trees[' + this.o_root.n_id + '].toggle(' + this.n_id + ')" onmouseover="trees[' + this.o_root.n_id + '].mover(' + this.n_id + ')" onmouseout="trees[' + this.o_root.n_id + '].mout(' + this.n_id + ')"><img src="' + this.get_icon(true) + '" border="0" align="absbottom" name="j_img' + this.o_root.n_id + '_' + this.n_id + '"></a>'
		: '<img src="' + this.get_icon(true) + '" border="0" align="absbottom">') : '') 
		+ '<a href="' + this.a_config[1] + '" target="' + this.o_root.a_tpl['target'] + '" onclick="return trees[' + this.o_root.n_id + '].select(' + this.n_id + ')" ondblclick="trees[' + this.o_root.n_id + '].toggle(' + this.n_id + ')" onmouseover="trees[' + this.o_root.n_id + '].mover(' + this.n_id + ')" onmouseout="trees[' + this.o_root.n_id + '].mout(' + this.n_id + ')" class="t' + this.o_root.n_id + 'i" id="i_txt' + this.o_root.n_id + '_' + this.n_id + '"><img src="' + this.get_icon() + '" border="0" align="absbottom" name="i_img' + this.o_root.n_id + '_' + this.n_id + '" class="t' + this.o_root.n_id + 'im">' + this.a_config[0] + '</a></td></tr></table>' + (this.a_children.length ? '<div id="i_div' + this.o_root.n_id + '_' + this.n_id + '" style="display:none"></div>' : '');
}

function item_get_icon (b_junction) {
	return this.o_root.a_tpl['icon_' + ((this.n_depth ? 0 : 32) + (this.a_children.length ? 16 : 0) + (this.a_children.length && this.b_opened ? 8 : 0) + (!b_junction && this.o_root.o_selected == this ? 4 : 0) + (b_junction ? 2 : 0) + (b_junction && this.is_last() ? 1 : 0))];
}

var trees = [];

get_element = document.all ?
	function (s_id) { return document.all[s_id] } :
	function (s_id) { return document.getElementById(s_id) };


//CLASP
function tree_get_opened_items(tw) {
	var s = "";
	for(var i=0;i<tw.a_flat_nodes.length;i++) {
		var node = tw.a_flat_nodes[i].ref;
		if(node.b_opened)  {
			s+=i + ",";
		}
	}
	return s;
}
function tree_restore_opened_items(tw) {
	for(var i=0;i<tw.open_list.length;i++) {
		if(tw.open_list[i]!="") {			
			var node = tw.a_flat_nodes[tw.open_list[i]].ref;
			if(!node.b_opened) {
				//document.write(node.init());
				node.open();
			}
		}
	}
}

function tree_persis_treeview_status() {
	for(var x=0;x<trees.length;x++) {
		var fld = clasp.form.getField(trees[x].menu_name + "_Status");
		if(fld)
			fld.value = tree_get_opened_items(trees[x].o_root);
	}
	return true;
}
//CLASP