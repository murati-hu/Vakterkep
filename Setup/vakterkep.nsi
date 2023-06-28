;Vakt�rk�p �s VAkt�rk�p szerkeszt� alkalmaz�s
;MUr�ti �kos 2002. m�rc. 25

;Telep�t� tulajdons�gai
NAme "Vakt�rk�p �s Vakt�rk�p Szerkeszt�"
Caption "Vakt�rk�p �s Vakt�rk�p Szerkeszt� telep�t�"
CRCCheck off
;Icon nagymo.ico


Subcaption 1 " - Telep�tend� komponensek"
Subcaption 2 " - Telep�t�s helye"
Subcaption 3 " - F�jlok m�sol�sa"
Subcaption 4 " - Telep�t�s v�ge"

MiscButtonText "< Vissza" "K�vetkez� >" "Kil�p�s" "Bez�r"
InstallButtonText "Telep�t�s"
SpaceTexts "Sz�ks�ges lemezter�let:  " "Rendelkez�sre �ll:  "
Icon "icon.ico"
Enabledbitmap "en.bmp"
disabledbitmap "dis.bmp"
Autoclosewindow true

Outfile "vti.exe"

InstallDir $PROGRAMFILES\Vakterkep

Componenttext "Ez a program telep�teni fogja a Vakt�rk�pet �s Vakt�rk�p Szerkeszt�t." "" "K�rem v�lassza ki a telep�tend� komponenseket"

DirText "V�lassza ki azt a helyet ahova telep�teni szeretn� a Vakt�rk�p �s Vakt�rk�p Szerkeszt�t" "Telep�t�s helye:" "Tall�z..."
Completedtext "A Telep�t�s sikeresen befejez�d�tt."
ShowInstdetails nevershow


;V�laszthat� komponensek
Section "Vakt�rk�p"
	SetOutpath $INSTDIR
	CreateDirectory "$SMPROGRAMS\Vakt�rk�p"
	CreateShortcut "$SMPROGRAMS\Vakt�rk�p\Vakt�rk�p.lnk" "$INSTDIR\vakterkep.exe" "" "$INSTDIR\vakterkep.exe" 0
	File "vakterkep.exe"
	File "vakterkep.chm"
	File "vakterkep.ini"
	CreateShortCut "$SMPROGRAMS\Vakt�rk�p\Vakt�rk�p S�g�.lnk" "$INSTDIR\vakterkep.chm" "" "$INSTDIR\vakterkep.chm" 0
SectionEnd

Section "Vakt�rk�p Szerkeszt�"
	CreateShortCut "$SMPROGRAMS\Vakt�rk�p\Vakt�rk�p Szerkeszt�.lnk" "$INSTDIR\vakterkep.exe" "/szerk" "$INSTDIR\vakterkep.exe" 0
	file "szerkeszto.chm"
	createShortCut "$SMPROGRAMS\Vakt�rk�p\Vakt�rk�p Szerkeszt� S�g�.lnk" "$INSTDIR\szerkeszto.chm" "" "$INSTDIR\szerkeszto.chm" 0
SectionEnd

Section "Be�ll�t�sok Enged�lyez�se"
	writeinistr "$INSTDIR\vakterkep.ini" "Telep�t� be�ll�t�sok" "beallitas" "1"	
SectionEnd

Section "T�rs�t�s"
	writeregstr "HKCR" ".vtk" "" "Vakt�rk�p f�jlok"
	writeregstr "HKCR" ".vtk\DefaultIcon" "" "$INSTDIR\vakterkep.exe,0"
	writeregstr "HKCR" ".vtk\shell" "" "Open"
	writeregstr "HKCR" ".vtk\shell\Open" "" ""
	writeregstr "HKCR" ".vtk\shell\open\command" "" '$INSTDIR\vakterkep.exe "%1"'
	
Sectionend

Section "Elt�vol�t�"
	WriteUninstaller "eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\Vakt�rk�p\Vakt�rk�p elt�vol�t�sa.lnk" "$INSTDIR\eltavolit.exe" "" "$INSTDIR\eltavolit.exe" 0
SectionEnd

;ELt�vol�t�
UninstallText "Ez az alkalmaz�s el fogja t�vol�tani a Vakt�rk�p �s Vakt�rk�p Szerkeszt�t." "Helye:"
UninstallCaption "Vakt�rk�p �s Vakt�rk�p Szerkeszt� elt�volol�t�sa"
UninstallButtonText "Elt�vol�t�s"
UninstallSUbCAption 0 " "
UninstallSUbCAption 1 " "
UninstallSUbCAption 2 " "
ShowUninstdetails nevershow

SEction "Uninstall"
	deleteregkey "HKCR" ".vtk" ""
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\Vakt�rk�p\*.*"
	rmdir $SMPROGRAMS\Vakt�rk�p
	rmdir $INSTDIR

Sectionend

