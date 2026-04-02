/*********************************************************/
/*                    Classe MENU                        */
/*********************************************************/

function PosicionaDiv(div, width, height, top, left) {
	var m = document.getElementById(div);	
	if (m != null) {
	    m.style.top = top + "px";
		m.style.left = left + "px";
		m.style.width = width + "px";
		m.style.height = (height + 12) + "px";
	}
}

function SetVisibleMenu() {
    var objTop = $("#" + this.obj).offset().top + $("#" + this.obj).height() + 5;
    var objLeft = $("#" + this.obj).offset().left - 13;

	var m = document.getElementById(this.div_menu);
	if (m != null) {
		if (this.exibe) {
		    PosicionaDiv(this.div_menu, this.width, this.height, objTop, objLeft);
		    m.style.display = "block";
		    this.exibe = false;
		} else {
			m.style.display = "none";
			this.exibe = true;
		}
	}
	this.ultimo_link = true;
}

function EscondeMenu() {
	var m = document.getElementById(this.div_menu);
	if (m != null) {
		m.style.display = "none";
	}
	if (!this.ultimo_link) {
		this.exibe = true;
	}
	this.ultimo_link = false;
}

function ClasseMenu(div_menu,width,height,obj) {
	this.exibe = true;
	this.ultimo_link = false;
	this.div_menu = div_menu;
	this.width = width;
	this.height = height;
	this.obj = obj;
	//Chamar no evento OnClick do body
	this.EsconderMenu = EscondeMenu;
	//Chamar no click no link que exibirá o menu
	this.SetarVisible = SetVisibleMenu;
}