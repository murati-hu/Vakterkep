;Vaktérkép és VAktérkép szerkesztõ alkalmazás
;MUráti Ákos 2002. márc. 25

;Telepítõ tulajdonságai
NAme "Vaktérkép és Vaktérkép Szerkesztõ"
Caption "Vaktérkép és Vaktérkép Szerkesztõ telepítõ"
CRCCheck off
;Icon nagymo.ico


Subcaption 1 " - Telepítendõ komponensek"
Subcaption 2 " - Telepítés helye"
Subcaption 3 " - Fájlok másolása"
Subcaption 4 " - Telepítés vége"

MiscButtonText "< Vissza" "Következõ >" "Kilépés" "Bezár"
InstallButtonText "Telepítés"
SpaceTexts "Szükséges lemezterület:  " "Rendelkezésre áll:  "
Icon "icon.ico"
Enabledbitmap "en.bmp"
disabledbitmap "dis.bmp"
Autoclosewindow true

Outfile "vti.exe"

InstallDir $PROGRAMFILES\Vakterkep

Componenttext "Ez a program telepíteni fogja a Vaktérképet és Vaktérkép Szerkesztõt." "" "Kérem válassza ki a telepítendõ komponenseket"

DirText "Válassza ki azt a helyet ahova telepíteni szeretné a Vaktérkép és Vaktérkép Szerkesztõt" "Telepítés helye:" "Tallóz..."
Completedtext "A Telepítés sikeresen befejezõdött."
ShowInstdetails nevershow


;Választható komponensek
Section "Vaktérkép"
	SetOutpath $INSTDIR
	CreateDirectory "$SMPROGRAMS\Vaktérkép"
	CreateShortcut "$SMPROGRAMS\Vaktérkép\Vaktérkép.lnk" "$INSTDIR\vakterkep.exe" "" "$INSTDIR\vakterkep.exe" 0
	File "vakterkep.exe"
	File "vakterkep.chm"
	File "vakterkep.ini"
	CreateShortCut "$SMPROGRAMS\Vaktérkép\Vaktérkép Súgó.lnk" "$INSTDIR\vakterkep.chm" "" "$INSTDIR\vakterkep.chm" 0
SectionEnd

Section "Vaktérkép Szerkesztõ"
	CreateShortCut "$SMPROGRAMS\Vaktérkép\Vaktérkép Szerkesztõ.lnk" "$INSTDIR\vakterkep.exe" "/szerk" "$INSTDIR\vakterkep.exe" 0
	file "szerkeszto.chm"
	createShortCut "$SMPROGRAMS\Vaktérkép\Vaktérkép Szerkesztõ Súgó.lnk" "$INSTDIR\szerkeszto.chm" "" "$INSTDIR\szerkeszto.chm" 0
SectionEnd

Section "Beállítások Engedélyezése"
	writeinistr "$INSTDIR\vakterkep.ini" "Telepítõ beállítások" "beallitas" "1"	
SectionEnd

Section "Társítás"
	writeregstr "HKCR" ".vtk" "" "Vaktérkép fájlok"
	writeregstr "HKCR" ".vtk\DefaultIcon" "" "$INSTDIR\vakterkep.exe,0"
	writeregstr "HKCR" ".vtk\shell" "" "Open"
	writeregstr "HKCR" ".vtk\shell\Open" "" ""
	writeregstr "HKCR" ".vtk\shell\open\command" "" '$INSTDIR\vakterkep.exe "%1"'
	
Sectionend

Section "Eltávolító"
	WriteUninstaller "eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\Vaktérkép\Vaktérkép eltávolítása.lnk" "$INSTDIR\eltavolit.exe" "" "$INSTDIR\eltavolit.exe" 0
SectionEnd

;ELtávolító
UninstallText "Ez az alkalmazás el fogja távolítani a Vaktérkép és Vaktérkép Szerkesztõt." "Helye:"
UninstallCaption "Vaktérkép és Vaktérkép Szerkesztõ eltávololítása"
UninstallButtonText "Eltávolítás"
UninstallSUbCAption 0 " "
UninstallSUbCAption 1 " "
UninstallSUbCAption 2 " "
ShowUninstdetails nevershow

SEction "Uninstall"
	deleteregkey "HKCR" ".vtk" ""
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\Vaktérkép\*.*"
	rmdir $SMPROGRAMS\Vaktérkép
	rmdir $INSTDIR

Sectionend

