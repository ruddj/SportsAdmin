; This Installer uses NSIS Multi User from
; https://github.com/Drizin/NsisMultiUser

!addplugindir /x86-ansi "..\NsisMultiUser\Plugins\x86-ansi"
!addplugindir /x86-unicode "..\NsisMultiUser\x86-unicode"
!addincludedir "..\NsisMultiUser\Include\"

!include MUI2.nsh
!include UAC.nsh
!include NsisMultiUser.nsh
!include LogicLib.nsh
!include "..\NsisMultiUser\Demos\Common\Utils.nsh"

!define PRODUCT_NAME "Sports Administrator" ; name of the application as displayed to the user
!define VERSION "5.3.1" ; main version of the application (may be 0.1, alpha, beta, etc.)
!define PROGEXE "Sports.accdr" ; main application filename
!define COMPANY_NAME "Sports Administrator" ; company, used for registry tree hierarchy
!define PRODUCT_FOLDER "SportsAdmin"
!define PRODUCT_WEB_SITE "https://github.com/ruddj/SportsAdmin"
;!define PLATFORM "Win32"
!define MIN_WIN_VER "XP"
!define SINGLE_INSTANCE_ID "${COMPANY_NAME} ${PRODUCT_NAME} Unique ID" ; do not change this between program versions!
!define LICENSE_FILE "License.txt" ; license file, optional

; NsisMultiUser optional defines
!define MULTIUSER_INSTALLMODE_ALLOW_BOTH_INSTALLATIONS 1 ; value 0 is not supported - previous installation is not fully removed
!define MULTIUSER_INSTALLMODE_ALLOW_ELEVATION 1
!define MULTIUSER_INSTALLMODE_ALLOW_ELEVATION_IF_SILENT 0
!define MULTIUSER_INSTALLMODE_DEFAULT_ALLUSERS 1
;!define MULTIUSER_INSTALLMODE_APPDATA 1    ; Install to ProgramData

!define MULTIUSER_INSTALLMODE_DISPLAYNAME "${PRODUCT_NAME} ${VERSION}"  

Var StartMenuFolder
Var AccessExe

; Installer Attributes
Name "${PRODUCT_NAME} ${VERSION}"
; The file to write
OutFile "${PRODUCT_FOLDER}-${VERSION}-Install.exe"
BrandingText "©2017 ${COMPANY_NAME}"

; SetShellVarContext all

AllowSkipFiles off
SetOverwrite on ; (default setting) set to on except for where it is manually switched off
ShowInstDetails show 
SetCompressor /SOLID lzma

; Pages
!define MUI_ABORTWARNING ; Show a confirmation when cancelling the installation

!define MUI_PAGE_CUSTOMFUNCTION_PRE PageWelcomeLicensePre
!insertmacro MUI_PAGE_WELCOME

!ifdef LICENSE_FILE
	!define MUI_PAGE_CUSTOMFUNCTION_PRE PageWelcomeLicensePre
	!insertmacro MUI_PAGE_LICENSE ".\${LICENSE_FILE}"
!endif

;!define MUI_PAGE_CUSTOMFUNCTION_PRE PageWelcomeLicensePre
;!insertmacro MUI_PAGE_LICENSE "README.md"

!define MULTIUSER_INSTALLMODE_CHANGE_MODE_FUNCTION PageInstallModeChangeMode
!insertmacro MULTIUSER_PAGE_INSTALLMODE 

!define MUI_COMPONENTSPAGE_SMALLDESC
!insertmacro MUI_PAGE_COMPONENTS

!define MUI_PAGE_CUSTOMFUNCTION_PRE PageDirectoryPre
!define MUI_PAGE_CUSTOMFUNCTION_SHOW PageDirectoryShow
!insertmacro MUI_PAGE_DIRECTORY

!define MUI_STARTMENUPAGE_NODISABLE ; Do not display the checkbox to disable the creation of Start Menu shortcuts
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "${PRODUCT_NAME}"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "SHCTX" ; writing to $StartMenuFolder happens in MUI_STARTMENU_WRITE_END, so it's safe to use "SHCTX" here
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "StartMenuFolder"
!define MUI_PAGE_CUSTOMFUNCTION_PRE PageStartMenuPre
!insertmacro MUI_PAGE_STARTMENU "" "$StartMenuFolder"

!define MUI_FINISHPAGE_NOAUTOCLOSE  ; Show Details Page
!insertmacro MUI_PAGE_INSTFILES

!define MUI_FINISHPAGE_RUN 
!define MUI_FINISHPAGE_RUN_FUNCTION PageFinishRun
!insertmacro MUI_PAGE_FINISH

!include Uninstall.nsh


!insertmacro MUI_LANGUAGE "English" ; Set languages (first is default language) - must be inserted after all pages 

InstType "Typical" 
InstType "Minimal" 
InstType "Full" 

Section "Core Files (required)" SectionCoreFiles
	SectionIn 1 2 3 RO
				
	${if} $HasCurrentModeInstallation == 1 ; if there's an installed version, remove all optional components (except "Core Files")
		; Clean up "Documentation"
		;Delete "$INSTDIR\readme.txt"
	
		; Clean up "Program Group" - we check that we created Start menu folder, if $StartMenuFolder is empty, the whole $SMPROGRAMS directory will be removed!
		${if} "$StartMenuFolder" != ""
			RMDir /r "$SMPROGRAMS\$StartMenuFolder"
		${endif}	
	
		; Clean up "Dektop Icon"
		Delete "$DESKTOP\${PRODUCT_NAME}.lnk"
		
		; Clean up "Start Menu Icon"
		Delete "$STARTMENU\${PRODUCT_NAME}.lnk"
			
		; Clean up "Quick Launch Icon"
		Delete "$QUICKLAUNCH\${PRODUCT_NAME}.lnk"		
	${endif}

	SetOutPath $INSTDIR
	; Write uninstaller and registry uninstall info as the first step,
	; so that the user has the option to run the uninstaller if sth. goes wrong 
	WriteUninstaller "${UNINSTALL_FILENAME}"			
	!insertmacro MULTIUSER_RegistryAddInstallInfo ; add registry keys		
	
     ; Put file there
	 ; Backup old file
	SetOverwrite on
	IfFileExists $INSTDIR\Sports.accdr 0 +3
	Delete $INSTDIR\Sports-old.accdr
	Rename $INSTDIR\Sports.accdr $INSTDIR\Sports-old.accdr
	 
	File /oname=Sports.accdr Sports.accdb
	File /oname=SportsAdmin.chm Source\help\SportsAdmin.chm
	File /oname=SportsAdmin.chw Source\help\SportsAdmin.chw   ; Adds Search to Help
	File Sports.ico
	File /oname=sports2.ico Source\installs\sports2.ico
	
	; Local Documentation
	File CHANGELOG.md
	File README.md
	File License.txt
	
	${if} $MultiUser.InstallMode == "AllUsers"
		; Allow Folder Security User Modify
		AccessControl::GrantOnFile "$INSTDIR" "(BU)" "GenericRead + GenericWrite + GenericExecute + Delete"
		Pop $R0
		${If} $R0 == error
			Pop $R0
			DetailPrint `AccessControl error: $R0`
		${EndIf}
	${endif}
	
SectionEnd

Section "HTML Templates" SectionTemplates	
	SectionIn 1 3
	
	SetOutPath $INSTDIR	
	File /nonfatal /a /r "web" 
	
SectionEnd

Section "Sample Databases" SectionDemo	
	SectionIn 1 3
	
	SetOutPath $INSTDIR	
	File /nonfatal /a /r "carnival"
	
SectionEnd

;Section "Sports View" SectionViewer
;	SectionIn 1 3
;	
;	SetOutPath $INSTDIR	
;	;File /nonfatal /a /r "demo"
;	
;SectionEnd

SectionGroup /e "Integration" SectionGroupShortcuts

Section "Program Group" SectionProgramGroup
	SectionIn 1	3
	
	!insertmacro MUI_STARTMENU_WRITE_BEGIN ""

	CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
	CreateShortcut "$SMPROGRAMS\$StartMenuFolder\${PRODUCT_NAME}.lnk"  \
	  $AccessExe "/runtime $\"$INSTDIR\${PROGEXE}$\"" "$INSTDIR\Sports.ico" 0
	  
	CreateShortcut "$SMPROGRAMS\$StartMenuFolder\${PRODUCT_NAME} Help.lnk"  \
	  "$INSTDIR\SportsAdmin.chm" "" 
	CreateDirectory "$SMPROGRAMS\$StartMenuFolder\Utilities"  
  	CreateShortcut "$SMPROGRAMS\$StartMenuFolder\Utilities\Compact ${PRODUCT_NAME}.lnk"  \
	  $AccessExe "/runtime /compact $\"$INSTDIR\${PROGEXE}$\""
	  
	CreateShortcut "$SMPROGRAMS\$StartMenuFolder\Utilities\Web Templates.lnk"  "$INSTDIR\web" 

		${if} $MultiUser.InstallMode == "AllUsers" 
			CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\${UNINSTALL_FILENAME}" "/allusers"
		${else}
			CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall (current user).lnk" "$INSTDIR\${UNINSTALL_FILENAME}" "/currentuser"
		${endif}
	
  !insertmacro MUI_STARTMENU_WRITE_END	
SectionEnd

Section  /o "Desktop Icon" SectionDesktopIcon
	SectionIn 3

	CreateShortcut "$DESKTOP\${PRODUCT_NAME}.lnk"  \
	  $AccessExe "/runtime $\"$INSTDIR\${PROGEXE}$\"" "$INSTDIR\Sports.ico" 0
	  
SectionEnd

Section /o "Start Menu Icon" SectionStartMenuIcon
	SectionIn 3

	CreateShortcut "$STARTMENU\${PRODUCT_NAME}.lnk"  \
	  $AccessExe "/runtime $\"$INSTDIR\${PROGEXE}$\"" "$INSTDIR\Sports.ico" 0
	  
SectionEnd

SectionGroupEnd

; Modern install component descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionCoreFiles} "Core files requred to run ${PRODUCT_NAME}."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionTemplates} "HTML Template files for ${PRODUCT_NAME}."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionDemo} "Sample Carnival Database files for ${PRODUCT_NAME}."
	;!insertmacro MUI_DESCRIPTION_TEXT ${SectionViewer} "Read-only version for competitors to view results."
	
  !insertmacro MUI_DESCRIPTION_TEXT ${SectionGroupShortcuts} "Select where to create shortcuts."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionProgramGroup} "Create a ${PRODUCT_NAME} program group under Start Menu->Programs."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionDesktopIcon} "Create ${PRODUCT_NAME} icon on the Desktop."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionStartMenuIcon} "Create ${PRODUCT_NAME} icon in the Start Menu."
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; Callbacks 
Function .onInit
	!insertmacro CheckMinWinVer ${MIN_WIN_VER}
	${ifnot} ${UAC_IsInnerInstance}
		!insertmacro CheckSingleInstance "${SINGLE_INSTANCE_ID}"
	${endif}	
	
	; Find Access
	; Need to read in location of MS Access from registry
	Call AccessLocation
	Pop $AccessExe
	
	${if} $AccessExe == ""
		MessageBox MB_ICONSTOP "Microsoft Access or Access Runtime were not found." /SD IDOK	
		SetErrorLevel ${MULTIUSER_ERROR_INVALID_PARAMETERS}
		Quit
	${endif}	
	
	!insertmacro MULTIUSER_INIT	  
FunctionEnd

Function PageWelcomeLicensePre		
	${if} $InstallShowPagesBeforeComponents == 0
		Abort ; don't display the Welcome and License pages for the inner instance 
	${endif}	
FunctionEnd

Function PageInstallModeChangeMode
	!insertmacro MUI_STARTMENU_GETFOLDER "" $StartMenuFolder
	 ; Install to ProgramData
	${if} $MultiUser.InstallMode == "AllUsers"
		StrCpy $INSTDIR "$APPDATA\${MULTIUSER_INSTALLMODE_INSTDIR}"
	${endif}
FunctionEnd

Function PageDirectoryPre	
	GetDlgItem $0 $HWNDPARENT 1		
	${if} ${SectionIsSelected} ${SectionProgramGroup}		
		SendMessage $0 ${WM_SETTEXT} 0 "STR:$(^NextBtn)" ; this is not the last page before installing
	${else}
		SendMessage $0 ${WM_SETTEXT} 0 "STR:$(^InstallBtn)" ; this is the last page before installing
	${endif}		
FunctionEnd

Function PageDirectoryShow
	${if} $CmdLineDir != ""
		${orif} $HasCurrentModeInstallation == 1
		FindWindow $R1 "#32770" "" $HWNDPARENT
		
		GetDlgItem $0 $R1 1019 ; Directory edit
		SendMessage $0 ${EM_SETREADONLY} 1 0 ; read-only is better than disabled, as user can copy contents
		
		GetDlgItem $0 $R1 1001 ; Browse button
		EnableWindow $0 0	
	${endif}			
FunctionEnd

Function PageStartMenuPre
	${ifnot} ${SectionIsSelected} ${SectionProgramGroup}
		Abort ; don't display this dialog if SectionProgramGroup is not selected
	${endif}	
FunctionEnd

Function PageFinishRun
	; the installer might exit too soon before the application starts and it loses the right to be the foreground window and starts in the background
	; however, if there's no active window when the application starts, it will become the active window, so we hide the installer
	HideWindow
	; the installer will show itself again quickly before closing (w/o Taskbar button), we move it offscreen
	!define SWP_NOSIZE 0x0001
	!define SWP_NOZORDER 0x0004
	System::Call "User32::SetWindowPos(i, i, i, i, i, i, i) b ($HWNDPARENT, 0, -1000, -1000, 0, 0, ${SWP_NOZORDER}|${SWP_NOSIZE})"

	!insertmacro UAC_AsUser_ExecShell "open" "$AccessExe" "/runtime $\"$INSTDIR\${PROGEXE}$\""  "$INSTDIR" ""
FunctionEnd

Function .onInstFailed
	MessageBox MB_ICONSTOP "${PRODUCT_NAME} ${VERSION} could not be fully installed.$\r$\nPlease, restart Windows and run the setup program again." /SD IDOK
FunctionEnd


;
; Returns the MS Access runtime executable by finding the correct registry key
;
;  Input: None
; Output: <Path to MS Access>
;
; Usage:
;
;    Call AccessVersion
;    Pop "$1"
;    MessageBox MB_OK|MB_ICONINFORMATION "Access version: $1"
;
 
Function AccessLocation
 
  ; Save R0,R1 on the stack
  Push $R1
  Push $R0
 
  ClearErrors
  ; Check a file association exists
  ReadRegStr $R0 HKEY_CLASSES_ROOT ".accdr" ""
  IfErrors NoModernAccess
  
  ; Read Access.ACCDRFile.16 to find Path
  ReadRegStr $R0 HKEY_CLASSES_ROOT "$R0\shell\Open\command" ""

  ; Search for Access 2016
  Push $R0  ; Input string
  Call GetExePart
  Pop $R0
  IfErrors NotFound
  StrCpy $R1 $R0
  Goto Found

  NoModernAccess:
	DetailPrint "Access Registry key not found: $0"
 
  NotFound:
    ; MessageBox MB_OK|MB_ICONEXCLAMATION "NSIS was not able to detect your MS Access version"
    StrCpy $R1 ""
 
  Found:
  Pop $R0
  Exch $R1
 
FunctionEnd

Function GetExePart
  Exch $R0
  Push $R1
  Push $R2
  StrLen $R1 $R0
  IntOp $R1 $R1 + 1
  loop:
    IntOp $R1 $R1 - 1
    StrCpy $R2 $R0 1 -$R1
    StrCmp $R2 "" exit2
    StrCmp $R2 "." exit1 ; Change " " to "\" if ur inputting dir path str
  Goto loop
  exit1:
    IntOp $R1 $R1 - 5  ; Move to 4 after character found
	StrCpy $R0 $R0 -$R1
	DetailPrint "Found Access at: $R0"
  exit2:
    Pop $R2
    Pop $R1
    Exch $R0
FunctionEnd