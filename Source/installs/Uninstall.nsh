Var RunningFromInstaller

; Installer Attributes
ShowUninstDetails show 

; Pages
!define MUI_UNABORTWARNING ; Show a confirmation when cancelling the installation

!define MULTIUSER_INSTALLMODE_CHANGE_MODE_UNFUNCTION un.PageInstallModeChangeMode
!insertmacro MULTIUSER_UNPAGE_INSTALLMODE

!define MUI_PAGE_CUSTOMFUNCTION_SHOW un.PageComponentsShow
!insertmacro MUI_UNPAGE_COMPONENTS

!insertmacro MUI_UNPAGE_INSTFILES

Section "un.Program Files" SectionUninstallProgram
	SectionIn RO

	; Try to delete the EXE as the first step - if it's in use, don't remove anything else
	!insertmacro un.DeleteRetryAbort "$INSTDIR\${PROGEXE}"
	!ifdef LICENSE_FILE
		!insertmacro un.DeleteRetryAbort "$INSTDIR\${LICENSE_FILE}"
	!endif
	
	Delete "$INSTDIR\*.ico"
	Delete "$INSTDIR\SportsAdmin.chm"
	; Clean up "Documentation"
	;!insertmacro un.DeleteRetryAbort "$INSTDIR\readme.txt"
	Delete "$INSTDIR\web\sample\*.*"

	
  ; Clean up "Program Group" - we check that we created Start menu folder, if $StartMenuFolder is empty, the whole $SMPROGRAMS directory will be removed!
	${if} "$StartMenuFolder" != ""
		RMDir /r "$SMPROGRAMS\$StartMenuFolder"
	${endif}	
	
  ; Clean up "Dektop Icon"
	!insertmacro un.DeleteRetryAbort "$DESKTOP\${PRODUCT_NAME}.lnk"
	
  ; Clean up "Start Menu Icon"
	!insertmacro un.DeleteRetryAbort "$STARTMENU\${PRODUCT_NAME}.lnk"
		
  ; Clean up "Quick Launch Icon"
	!insertmacro un.DeleteRetryAbort "$QUICKLAUNCH\${PRODUCT_NAME}.lnk"	
SectionEnd

Section /o "un.Program Settings" SectionRemoveSettings
  ; this section is executed only explicitly and shouldn't be placed in SectionUninstallProgram
	DeleteRegKey HKCU "Software\${PRODUCT_NAME}"	
	Delete "$INSTDIR\carnival\*.*"
	RMDir /r "$INSTDIR\carnival"	
	
    Delete "$INSTDIR\web\*.*"
	RMDir /r "$INSTDIR\web"
SectionEnd

Section "-Uninstall" ; hidden section, must always be the last one!
	; Remove the uninstaller from registry as the very last step - if sth. goes wrong, let the user run it again
	!insertmacro MULTIUSER_RegistryRemoveInstallInfo ; Remove registry keys
		
  Delete "$INSTDIR\${UNINSTALL_FILENAME}"	
  ; remove the directory only if it is empty - the user might have saved some files in it		
	RMDir "$INSTDIR"  		
SectionEnd

; Modern install component descriptions
!insertmacro MUI_UNFUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionUninstallProgram} "Uninstall ${PRODUCT_NAME} files."
	!insertmacro MUI_DESCRIPTION_TEXT ${SectionRemoveSettings} "Remove ${PRODUCT_NAME} program data. Select only if you don't plan to use the program in the future."
!insertmacro MUI_UNFUNCTION_DESCRIPTION_END

; Callbacks
Function un.onInit
	${GetParameters} $R0
		
	${GetOptions} $R0 "/uninstall" $R1
	${ifnot} ${errors}	
		StrCpy $RunningFromInstaller 1		
	${else}
		StrCpy $RunningFromInstaller 0
	${endif}
	
	${ifnot} ${UAC_IsInnerInstance}
		${andif} $RunningFromInstaller == "0"
		!insertmacro CheckSingleInstance "${SINGLE_INSTANCE_ID}"
	${endif}		
		
	!insertmacro MULTIUSER_UNINIT		
FunctionEnd

Function un.PageInstallModeChangeMode
	!insertmacro MUI_STARTMENU_GETFOLDER "" $StartMenuFolder
FunctionEnd

Function un.PageComponentsShow
	; Show/hide the Back button 
	GetDlgItem $0 $HWNDPARENT 3 
	ShowWindow $0 $UninstallShowBackButton
FunctionEnd

Function un.onUninstFailed
	MessageBox MB_ICONSTOP "${PRODUCT_NAME} ${VERSION} could not be fully uninstalled.$\r$\nPlease, restart Windows and run the uninstaller again." /SD IDOK	
FunctionEnd
