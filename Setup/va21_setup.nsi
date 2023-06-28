;Vakablak 2.1 telepítõ

!include "MUI.nsh"

!define M "Vakablak"
!define SZERK "Vakablak Szerkesztõ 2.1"
!define VA "Vakablak 2.1"

;--------------------------------
;Konfig:
	Name "${VA}"
	OutFile "vas21.exe"
	ShowInstDetails show

	InstallDir "$PROGRAMFILES\${M}"
  	InstallDirRegKey HKCU "Software\${M}" ""

;--------------------------------
;Modern UI Configuration

  !insertmacro MUI_PAGE_LICENSE "license.rtf"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"
  
;--------------------------------
;Language Strings

  ;Description
	LangString DESC_Vakablak ${LANG_HUNGARIAN} "A Vakablak telepítése az ön számítógépére."
	LangString DESC_Szerkeszto ${LANG_HUNGARIAN} "Szerkesztõ telepítése az ön számítógépére"
	LangString DESC_Nyelvek ${LANG_HUNGARIAN} "Más nyelvek telepítése: English, Deutsch, Francais"
	LangString DESC_Tarsitas ${LANG_HUNGARIAN} "Projekt fájlok társítása az alkalmazásokhoz"
	LangString DESC_VB6 ${LANG_HUNGARIAN} "A futtatáshoz szükséges Visual Basic 6.0 (SP5) Runtime fájlok telepítése.(XP alatt nem szükséges)"
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Eltávolító alkalmazás telepítése (Uninstall)"
;--------------------------------

;Installer Sections

Section "${VA}" Vakablak
	SectionIn RO
	
	;*****************************
	detailprint ">>> Korábban telepített komponensek eltávolítása..."
	
	;Régi Szerkesztõ törlése:
		delete "$SMPROGRAMS\${M}\${SZERK}.lnk"

	;Súgó törlése
		delete "$SMPROGRAMS\${M}\${VA} Súgó.lnk"
		delete "$INSTDIR\vakablak.chm"

	;Szerk súgó törlése
		delete "$SMPROGRAMS\${M}\${SZERK} Súgó.lnk"
		delete "$INSTDIR\szerkeszto.chm"

	;Társítás törlése
		deleteregkey "HKCR" ".vtk" ""

	;Eltávolító törlése
		delete "$INSTDIR\*.*"
		delete "$SMPROGRAMS\${M}\Eltávolítás.lnk"
		
	detailprint ""
	;*****************************

	WriteRegStr HKCU "Software\${M}" "" $INSTDIR

	detailprint ">>> Microsoft Commondialog ActiveX vezérlõ telepítése..."
	setoutpath $SYSDIR
	file "comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""
	
	detailprint ">>> Microsoft Standard Data Formating Object DLL telepítése..."
	file "msstdfmt.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/msstdfmt.dll"	
	detailprint ""

	detailprint ">>> ${VA} telepítése..."
  	SetOutPath "$INSTDIR"
	File "..\forras\vakablak.exe"
  	File "..\forras\ertekeles.ini"
	CreateDirectory "$SMPROGRAMS\${M}"
	CreateShortCut "$SMPROGRAMS\${M}\${VA}.lnk" "$INSTDIR\vakablak.exe"

	file "vakablak.url"
	CreateShortCut "$SMPROGRAMS\${M}\${VA} az interneten.lnk" "$INSTDIR\vakablak.url"

	writeregstr "HKCR" "Vakablak" "" "Vakablak Projekt"
	writeregstr "HKCR" "Vakablak\DefaultIcon" "" "$INSTDIR\vakablak.ico,0"
	writeregstr "HKCR" "Vakablak\shell" "" "Open"
	writeregstr "HKCR" "vakablak\shell\Open" "" ""
	writeregstr "HKCR" "vakablak\shell\open\command" "" "$INSTDIR\vakablak.exe %1"
	detailprint ""
SectionEnd

Section "${SZERK}" Szerkeszto
	detailprint ">>> ${SZERK} telepítése..."
	file "vakablak.ico"
	CreateShortCut "$SMPROGRAMS\${M}\${SZERK}.lnk" "$INSTDIR\vakablak.exe" "-sz" "$INSTDIR\vakablak.ico" 0
	
	writeregstr "HKCR" "Vakablak\shell\" "" "Open"
	writeregstr "HKCR" "Vakablak\shell\edit" "" "&Vakablak Szerkesztõ"
	writeregstr "HKCR" "Vakablak\shell\edit\command" "" "$INSTDIR\vakablak.exe -sz=%1"
	writeregstr "HKCR" "Vakablak\shell\kezi" "" "&Szerkesztés Jegyzettömbbel"
	writeregstr "HKCR" "Vakablak\shell\kezi\command" "" "notepad.exe %1"
	detailprint ""
SectionEnd

Section "Idegen nyelvek" Nyelvek
	detailprint ">>> Nyelvek másolása..."
  	SetOutPath "$INSTDIR"
	CreateDirectory "$INSTDIR\nyelvek\"
	SetOutPath "$INSTDIR\nyelvek\"	
	File "..\forras\nyelvek\*.lng"
SectionEnd

SubSection "Súgó"
	Section "${VA} Súgó"
		detailprint ">>> ${VA} Súgó telepítése..."
		SetOutPath "$INSTDIR"
		File "..\help\vakablak\vakablak.chm"
		CreateShortCut "$SMPROGRAMS\${M}\${VA} Súgó.lnk" "$INSTDIR\vakablak.chm"
		detailprint ""
	sectionend
	
	Section "${SZERK} Súgó"
		detailprint ">>> ${SZERK} Súgó telepítése..."
		SetOutPath "$INSTDIR"
		file "..\help\szerkeszto\szerkeszto.chm"
		createShortCut "$SMPROGRAMS\${M}\${SZERK} Súgó.lnk" "$INSTDIR\szerkeszto.chm" "" "$INSTDIR\szerkeszto.chm" 0
		detailprint ""
	sectionend
Subsectionend


Section "Projektek társítása" Tarsitas
	detailprint ">>> Fájlok társítása..."
	writeregstr "HKCR" ".vtk" "" "Vakablak"
	detailprint ""
Sectionend

section "Microsoft Visual Basic 6.0 Runtime (SP5)" VB6
	detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telepítése..."
	setoutpath $SYSDIR
	file "vbrun.exe"
	execwait "$SYSDIR\vbrun.exe /q"
	detailprint ""
sectionend

Section "Eltávolító alkalmazás" Eltavolit
	detailprint ">>> Eltávoító alkalmazás telepítése..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\${M}\Eltávolítás.lnk" "$INSTDIR\eltavolit.exe" 
Sectionend


!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${Vakablak} $(DESC_Vakablak)
	!insertmacro MUI_DESCRIPTION_TEXT ${Szerkeszto} $(DESC_Szerkeszto)
	!insertmacro MUI_DESCRIPTION_TEXT ${Nyelvek} $(DESC_Nyelvek)
	!insertmacro MUI_DESCRIPTION_TEXT ${Tarsitas} $(DESC_Tarsitas)
	!insertmacro MUI_DESCRIPTION_TEXT ${VB6} $(DESC_VB6)
	!insertmacro MUI_DESCRIPTION_TEXT ${Eltavolit} $(DESC_Eltavolit)
!insertmacro MUI_FUNCTION_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"
	deleteregkey "HKCR" ".vtk" ""
	deleteregkey "HKCR" "Vakablak" ""
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\${M}\*.*"
	rmdir "$SMPROGRAMS\${M}"
	rmdir "$INSTDIR"
SectionEnd