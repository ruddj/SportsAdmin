; SportAdminInstaller.nsi
;
; This script is based on example1.nsi, but it remember the directory, 
; has uninstall support and (optionally) installs start menu shortcuts.
;
; It will install example2.nsi into a directory that the user selects,

; Requires AccessControl plugin http://nsis.sourceforge.net/AccessControl_plug-in

;--------------------------------

!define PRODUCT_NAME "Sports Administrator"
!define VERSION 5.0
!define PRODUCT_VERSION 5.0
!define PRODUCT_GROUP "Sports Administrator"
!define PRODUCT_FOLDER "SportsAdmin"
!define PRODUCT_PUBLISHER "Sports Administrator"
!define PRODUCT_WEB_SITE "https://github.com/ruddj/SportsAdmin"
!define PRODUCT_DIR_REGKEY "Software\SportsAdmin"
!define CSIDL_COMMON_APPDATA 0x0023        
Var AppDataPath
; The name of the installer
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"

; The file to write
OutFile "sportadmin-${VERSION}-win32.exe"

; The default installation directory
InstallDir $PROGRAMFILES\${PRODUCT_FOLDER}

!ifdef NSIS_LZMA_COMPRESS_WHOLE
SetCompressor lzma
!else
SetCompressor /SOLID lzma
!endif

SetOverwrite ifnewer
CRCCheck on
BrandingText "${PRODUCT_NAME}"

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\${PRODUCT_FOLDER}" "Install_Dir"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

!include "StrFunc.nsh"

System::Call "shell32::SHGetFolderPath(0, i ${CSIDL_COMMON_APPDATA}, 0, 0, t .r1)"  
StrCpy $AppDataPath "$1\${PRODUCT_PUBLISHER}\${PRODUCT_NAME}"

;--------------------------------

; Pages

Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section "SportAdmin (required)"

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ;CreateDirectory "$AppDataPath"
  
  ; Put file there
  File /oname=Sports.accdr Sports.accdb
  File Sports.ico
  File sports2.ico
 ; Dir web
  File /nonfatal /a /r "web"
  
  ; Write the installation path into the registry
  WriteRegStr HKLM SOFTWARE\${PRODUCT_FOLDER} "Install_Dir" "$INSTDIR"
  
  ;AccessControl::GrantOnFile "$AppDataPath" "(BU)" "GenericRead + GenericWrite"
  
  ; Need to find Access Version and write Folder to Trusted Location
  
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_FOLDER}" "DisplayName" "${PRODUCT_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_FOLDER}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_FOLDER}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_FOLDER}" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

; Optional section (can be disabled by the user)
Section "Start Menu Shortcuts"


  CreateDirectory "$SMPROGRAMS\${PRODUCT_GROUP}"
  CreateShortcut "$SMPROGRAMS\${PRODUCT_GROUP}\Uninstall.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  ; Need to read in location of MS Access from registry
  Call AccessLocation
  Pop $R0
  CreateShortcut "$SMPROGRAMS\${PRODUCT_GROUP}\${PRODUCT_GROUP}.lnk" \
  $R0 "/runtime $\"$INSTDIR\Sports.accdr$\"" "$INSTDIR\Sports.ico" 0
  DetailPrint "CreateShortcut $SMPROGRAMS\${PRODUCT_GROUP}\${PRODUCT_GROUP}.lnk  $R0 /runtime $INSTDIR\Sports.accdr $INSTDIR\Sports.ico 0"

  
SectionEnd


;Section "Sample Database"

  ; Set output path to the installation directory.
;  SetOutPath $INSTDIR
  
  ; Put file there
;  File "example2.nsi"
  
;SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_FOLDER}"
  DeleteRegKey HKLM ${PRODUCT_DIR_REGKEY}

  ; Remove files and uninstaller
  Delete $INSTDIR\Sports.accdr
  Delete $INSTDIR\*.ico
  Delete $INSTDIR\uninstall.exe
  Delete $INSTDIR\web\sample\*.*
  Delete $INSTDIR\web\*.*


  ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\${PRODUCT_GROUP}\*.*"

  ; Remove directories used
  RMDir "$SMPROGRAMS\${PRODUCT_GROUP}"
  RMDir "$INSTDIR"

SectionEnd


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