;Vakablak 2.1 telep�t�

!include "MUI.nsh"

!define M "Vakablak"
!define SZERK "Vakablak Szerkeszt� 2.1"
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
	LangString DESC_Vakablak ${LANG_HUNGARIAN} "A Vakablak telep�t�se az �n sz�m�t�g�p�re."
	LangString DESC_Szerkeszto ${LANG_HUNGARIAN} "Szerkeszt� telep�t�se az �n sz�m�t�g�p�re"
	LangString DESC_Nyelvek ${LANG_HUNGARIAN} "M�s nyelvek telep�t�se: English, Deutsch, Francais"
	LangString DESC_Tarsitas ${LANG_HUNGARIAN} "Projekt f�jlok t�rs�t�sa az alkalmaz�sokhoz"
	LangString DESC_VB6 ${LANG_HUNGARIAN} "A futtat�shoz sz�ks�ges Visual Basic 6.0 (SP5) Runtime f�jlok telep�t�se.(XP alatt nem sz�ks�ges)"
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Elt�vol�t� alkalmaz�s telep�t�se (Uninstall)"
;--------------------------------

;Installer Sections

Section "${VA}" Vakablak
	SectionIn RO
	
	;*****************************
	detailprint ">>> Kor�bban telep�tett komponensek elt�vol�t�sa..."
	
	;R�gi Szerkeszt� t�rl�se:
		delete "$SMPROGRAMS\${M}\${SZERK}.lnk"

	;S�g� t�rl�se
		delete "$SMPROGRAMS\${M}\${VA} S�g�.lnk"
		delete "$INSTDIR\vakablak.chm"

	;Szerk s�g� t�rl�se
		delete "$SMPROGRAMS\${M}\${SZERK} S�g�.lnk"
		delete "$INSTDIR\szerkeszto.chm"

	;T�rs�t�s t�rl�se
		deleteregkey "HKCR" ".vtk" ""

	;Elt�vol�t� t�rl�se
		delete "$INSTDIR\*.*"
		delete "$SMPROGRAMS\${M}\Elt�vol�t�s.lnk"
		
	detailprint ""
	;*****************************

	WriteRegStr HKCU "Software\${M}" "" $INSTDIR

	detailprint ">>> Microsoft Commondialog ActiveX vez�rl� telep�t�se..."
	setoutpath $SYSDIR
	file "comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""
	
	detailprint ">>> Microsoft Standard Data Formating Object DLL telep�t�se..."
	file "msstdfmt.dll"
	execwait "regsvr32.exe /i /s $SYSDIR/msstdfmt.dll"	
	detailprint ""

	detailprint ">>> ${VA} telep�t�se..."
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
	detailprint ">>> ${SZERK} telep�t�se..."
	file "vakablak.ico"
	CreateShortCut "$SMPROGRAMS\${M}\${SZERK}.lnk" "$INSTDIR\vakablak.exe" "-sz" "$INSTDIR\vakablak.ico" 0
	
	writeregstr "HKCR" "Vakablak\shell\" "" "Open"
	writeregstr "HKCR" "Vakablak\shell\edit" "" "&Vakablak Szerkeszt�"
	writeregstr "HKCR" "Vakablak\shell\edit\command" "" "$INSTDIR\vakablak.exe -sz=%1"
	writeregstr "HKCR" "Vakablak\shell\kezi" "" "&Szerkeszt�s Jegyzett�mbbel"
	writeregstr "HKCR" "Vakablak\shell\kezi\command" "" "notepad.exe %1"
	detailprint ""
SectionEnd

Section "Idegen nyelvek" Nyelvek
	detailprint ">>> Nyelvek m�sol�sa..."
  	SetOutPath "$INSTDIR"
	CreateDirectory "$INSTDIR\nyelvek\"
	SetOutPath "$INSTDIR\nyelvek\"	
	File "..\forras\nyelvek\*.lng"
SectionEnd

SubSection "S�g�"
	Section "${VA} S�g�"
		detailprint ">>> ${VA} S�g� telep�t�se..."
		SetOutPath "$INSTDIR"
		File "..\help\vakablak\vakablak.chm"
		CreateShortCut "$SMPROGRAMS\${M}\${VA} S�g�.lnk" "$INSTDIR\vakablak.chm"
		detailprint ""
	sectionend
	
	Section "${SZERK} S�g�"
		detailprint ">>> ${SZERK} S�g� telep�t�se..."
		SetOutPath "$INSTDIR"
		file "..\help\szerkeszto\szerkeszto.chm"
		createShortCut "$SMPROGRAMS\${M}\${SZERK} S�g�.lnk" "$INSTDIR\szerkeszto.chm" "" "$INSTDIR\szerkeszto.chm" 0
		detailprint ""
	sectionend
Subsectionend


Section "Projektek t�rs�t�sa" Tarsitas
	detailprint ">>> F�jlok t�rs�t�sa..."
	writeregstr "HKCR" ".vtk" "" "Vakablak"
	detailprint ""
Sectionend

section "Microsoft Visual Basic 6.0 Runtime (SP5)" VB6
	detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telep�t�se..."
	setoutpath $SYSDIR
	file "vbrun.exe"
	execwait "$SYSDIR\vbrun.exe /q"
	detailprint ""
sectionend

Section "Elt�vol�t� alkalmaz�s" Eltavolit
	detailprint ">>> Elt�vo�t� alkalmaz�s telep�t�se..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\${M}\Elt�vol�t�s.lnk" "$INSTDIR\eltavolit.exe" 
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