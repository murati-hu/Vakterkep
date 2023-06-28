tablazat='<table bgcolor="#428AE7" border="0" align="center" width="80%"><tr>';

menupont("Kezdõlap","#","klikk('1');");
menupont("Vaktérkép Súgó","vt_frm.htm");
menupont("Vaktérkép Szerkesztõ Súgó","szerk_frm.htm");
menupont("Letöltések","#","klikk('2');");
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