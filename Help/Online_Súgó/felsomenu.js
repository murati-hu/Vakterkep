tablazat='<table bgcolor="#428AE7" border="0" align="center" width="80%"><tr>';

menupont("Kezd�lap","#","klikk('1');");
menupont("Vakt�rk�p S�g�","vt_frm.htm");
menupont("Vakt�rk�p Szerkeszt� S�g�","szerk_frm.htm");
menupont("Let�lt�sek","#","klikk('2');");
vege();













function menupont(szoveg, link, parancs) {	
	tablazat+='<td align="center" class="menupont"';
	if (parancs != '') {
			kieg=' onclick="' + parancs + '"';}
		else	{
			kieg=''; }
	tablazat+=kieg + '><a href="' + link + '"><b>' + szoveg + '</b></a></td>';
}

function vege() {
	tablazat+='</tr></table>';
	document.write(tablazat);

}

function klikk(idje) {
	for (i=1;i<=2;i++) {
		eval('tema'+ i+ '.style.display="none";');
	}
	eval('tema'+ idje + '.style.display="block";');


}