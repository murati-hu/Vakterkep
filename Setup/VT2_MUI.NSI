;Vaktérkép Moder telepítrõcsomag

!define MUI_PRODUCT "Vaktérkép" ;Define your own software name here
!define MUI_VERSION "2.0" ;Define your own software version here
!define VT "${MUI_PRODUCT} ${MUI_VERSION}"
!define SZERK "${MUI_PRODUCT} Szerkesztõ ${MUI_VERSION}"

!include "MUI.nsh"

;--------------------------------
;Configuration

  OutFile "vts2.exe"

  ShowInstDetails show

  InstallDir "$PROGRAMFILES\${VT}"
  
  InstallDirRegKey HKCU "Software\${VT}" ""

;--------------------------------
;Modern UI Configuration

  ;!define USE_SUBSECTION
  !define MUI_LICENSEPAGE
  !define MUI_COMPONENTSPAGE
  !define MUI_COMPONENTSPAGE_SMALLDESC
  !define MUI_DIRECTORYPAGE
  
  !define MUI_ABORTWARNING
  
  !define MUI_UNINSTALLER
  !define MUI_UNCONFIRMPAGE
  
  ;!define MUI_HEADERBITMAP "${NSISDIR}\Contrib\Icons\modern-header.bmp"
  ;!define MUI_ICON "${NSISDIR}\Contrib\Icons\setup.ico"
  ;!define MUI_UNICON "${NSISDIR}\Contrib\Icons\setup.ico"
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"
  
;--------------------------------
;Language Strings

  ;Description
	LangString DESC_Vakterkep ${LANG_HUNGARIAN} "${VT} telepítése az ön számítógépére."
	LangString DESC_Szerkeszto ${LANG_HUNGARIAN} "${SZERK} telepítése az ön számítógépére"
	LangString DESC_Tarsitas ${LANG_HUNGARIAN} "${VT} Projekt fájlok (*.vtk) társítása az alkalmazásokhoz."
	LangString DESC_VB6 ${LANG_HUNGARIAN} "A ${VT} futásához szükséges Visual Basic 6.0 (SP5) Runtime fájlok telepítése.(XP alatt nem szükséges)"
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "${VT} eltávolító alkalmazás telepítése. (Uninstall)"

;--------------------------------
;Data
  
  LicenseData "license.rtf"

;--------------------------------
;Reserve Files
  
  ;Things that need to be extracted on first (keep these lines before any File command!)
  ;Only useful for BZIP2 compression
  
  ;ReserveFile "${NSISDIR}\Contrib\Icons\modern-header.bmp"

;--------------------------------
;Installer Sections

Section "${VT}" Vakterkep
	SectionIn RO
	WriteRegStr HKCU "Software\${VT}" "" $INSTDIR

	detailprint ">>> Microsoft Commondialog ActiveX vezérlõ telepítése..."
	setoutpath $SYSDIR
	file "comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""
	
	detailprint ">>> Microsoft Standard Data Formating Object DLL telepítése..."
	file "msstdfmt.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/msstdfmt.dll"	
	detailprint ""

	detailprint ">>> ${VT} telepítése..."
  	SetOutPath "$INSTDIR"
	File "vakterkep.exe"
  	CreateDirectory "$SMPROGRAMS\${VT}"
	CreateShortCut "$SMPROGRAMS\${VT}\${VT}.lnk" "$INSTDIR\vakterkep.exe"

	writeregstr "HKCR" "Vakterkep" "" "Vaktérkép Projekt"
	writeregstr "HKCR" "Vakterkep\DefaultIcon" "" "$INSTDIR\vakterkep.ico,0"
	writeregstr "HKCR" "Vakterkep\shell" "" "Open"
	writeregstr "HKCR" "vakterkep\shell\Open" "" ""
	writeregstr "HKCR" "vakterkep\shell\open\command" "" "$INSTDIR\vakterkep.exe %1"
	detailprint ""
SectionEnd

Section "${SZERK}" Szerkeszto
	detailprint ">>> ${SZERK} telepítése..."
	file "vakterkep.ico"
	CreateShortCut "$SMPROGRAMS\${VT}\${SZERK}.lnk" "$INSTDIR\vakterkep.exe" "-sz" "$INSTDIR\vakterkep.ico" 0
	
	writeregstr "HKCR" "Vakterkep\shell\" "" "Open"
	writeregstr "HKCR" "Vakterkep\shell\edit" "" "&Szerkesztés"
	writeregstr "HKCR" "Vakterkep\shell\edit\command" "" "$INSTDIR\vakterkep.exe -sz=%1"
	writeregstr "HKCR" "Vakterkep\shell\kezi" "" "&Kézi szerkesztés Jegyzettömbbel"
	writeregstr "HKCR" "Vakterkep\shell\kezi\command" "" "notepad.exe %1"
	detailprint ""
SectionEnd


SubSection "Súgó"
	Section "${VT} Súgó"
		detailprint ">>> ${VT} Súgó telepítése..."
		File "vakterkep.chm"
		CreateShortCut "$SMPROGRAMS\${VT}\${VT} Súgó.lnk" "$INSTDIR\vakterkep.chm"
		detailprint ""
	sectionend
	
	Section "${SZERK} Súgó"
		detailprint ">>> ${SZERK} Súgó telepítése..."
		file "szerkeszto.chm"
		createShortCut "$SMPROGRAMS\${VT}\${SZERK} Súgó.lnk" "$INSTDIR\szerkeszto.chm" "" "$INSTDIR\szerkeszto.chm" 0
		detailprint ""
	sectionend
Subsectionend


Section "Térkép Fájlok Társítása" Tarsitas
	detailprint ">>> Fájlok társítása..."
	writeregstr "HKCR" ".vtk" "" "Vakterkep"
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
	CreateShortCut "$SMPROGRAMS\${VT}\Eltávolítás.lnk" "$INSTDIR\eltavolit.exe" 
Sectionend 


!insertmacro MUI_SECTIONS_FINISHHEADER


!insertmacro MUI_FUNCTIONS_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${Vakterkep} $(DESC_Vakterkep)
	!insertmacro MUI_DESCRIPTION_TEXT ${Szerkeszto} $(DESC_Szerkeszto)
	!insertmacro MUI_DESCRIPTION_TEXT ${Tarsitas} $(DESC_Tarsitas)
	!insertmacro MUI_DESCRIPTION_TEXT ${VB6} $(DESC_VB6)
	!insertmacro MUI_DESCRIPTION_TEXT ${Eltavolit} $(DESC_Eltavolit)
!insertmacro MUI_FUNCTIONS_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"
	deleteregkey "HKCR" ".vtk" ""
	delete "$INSTDIR\*.*"
	delete "$SMPROGRAMS\${VT}\*.*"
	rmdir "$SMPROGRAMS\${VT}"
	rmdir "$INSTDIR"

  	!insertmacro MUI_UNFINISHHEADER
SectionEnd