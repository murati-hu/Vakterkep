var globalis
function mutat(mit) {
	var allapot =""
	eval("allapot=" +  mit + ".style.display;");
	if (allapot == "none") {
			allapot="block"}
	else {
			allapot="none" }
	eval(mit + ".style.display='" + allapot + "';");

}

function latszat() {
	kiir="<hr><h6><sup>*</sup>Az itt látható képek nem feltétlenül így néznek ki így az ön számítógépén. Az ablakok szinei, kinézetei attól függ, hogy ön melyik Windows Operációs rendszert használja.</h6>"
	document.write(kiir);



}

function kapcsolat() {
	document.write('<p><b><address>Muráti Ákos<br>e-mail: <a href="mailto:b0murako@gyakg.u-szeged.hu">b0murako@gyakg.u-szeged.hu</a><br>honlap: <a href="http://www.tar.hu/murako">www.tar.hu/murako</a><br></address></b></p>');

}