
' This utility is by Tobias Kienzler from 
' http://stackoverflow.com/questions/187506/how-do-you-use-version-control-with-access-development/35308115

Option Explicit

const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please enter the filename!", vbExclamation, "Error"
    Wscript.Quit()
End if
sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sExportpath
If (WScript.Arguments.Count = 1) then
    sExportpath = ""
else
    sExportpath = WScript.Arguments(1)
End If


exportModulesTxt sADPFilename, sExportpath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function exportModulesTxt(sADPFilename, sExportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring
	Dim dbs, i
	Const acQuery = 1

    dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    If (sExportpath = "") then
        sExportpath = myPath & "\Source\"
    End If
    sStubADPFilename = sExportpath & myName & "_stub." & myType

    WScript.Echo "copy stub to " & sStubADPFilename & "..."
    On Error Resume Next
        fso.CreateFolder(sExportpath)
    On Error Goto 0
    fso.CopyFile sADPFilename, sStubADPFilename

    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sStubADPFilename & " ..."
    If (Right(sStubADPFilename,4) = ".adp") Then
        oApplication.OpenAccessProject sStubADPFilename
    Else
        oApplication.OpenCurrentDatabase sStubADPFilename
    End If

    oApplication.Visible = false

    dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."
    Dim myObj
    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acForm, myObj.fullname, sExportpath & "\" & strClean (myObj.fullname) & ".form"
        oApplication.DoCmd.Close acForm, myObj.fullname
        dctDelete.Add "FO" & myObj.fullname, acForm
    Next
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acModule, myObj.fullname, sExportpath & "\" & strClean (myObj.fullname) & ".bas"
        dctDelete.Add "MO" & myObj.fullname, acModule
    Next
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acMacro, myObj.fullname, sExportpath & "\" & strClean (myObj.fullname) & ".mac"
        dctDelete.Add "MA" & myObj.fullname, acMacro
    Next
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullname
        oApplication.SaveAsText acReport, myObj.fullname, sExportpath & "\" & strClean (myObj.fullname) & ".report"
        dctDelete.Add "RE" & myObj.fullname, acReport
    Next
	Set dbs = oApplication.CurrentDb() 
	For i = 0 To dbs.QueryDefs.Count - 1
	    WScript.Echo "  " & dbs.QueryDefs(i).Name
        oApplication.SaveAsText acQuery, dbs.QueryDefs(i).Name,  sExportpath & "\" & strClean (dbs.QueryDefs(i).Name )& ".txt"
        'dctDelete.Add "qry_" & dbs.QueryDefs(i).Name, acQuery
    Next

    WScript.Echo "deleting..."
    dim sObjectname
    For Each sObjectname In dctDelete
        WScript.Echo "  " & Mid(sObjectname, 3)
        oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
    Next

    oApplication.CloseCurrentDatabase
    oApplication.CompactRepair sStubADPFilename, sStubADPFilename & "_"
    oApplication.Quit

    fso.CopyFile sStubADPFilename & "_", sStubADPFilename
    fso.DeleteFile sStubADPFilename & "_"


End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function

Function strClean (strtoclean)
	Dim objRegExp, outputStr
	Set objRegExp = New Regexp

	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[(?*"",\\<>&#~%{}+.@:\/!;]+"
	outputStr = objRegExp.Replace(strtoclean, "-")

	objRegExp.Pattern = "\-+"
	outputStr = objRegExp.Replace(outputStr, "-")

	strClean = outputStr
End Function