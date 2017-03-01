Option Compare Database

'################################################################
'#                                                              #
'#      Created with / Erstellt mit:                            #
'#      IDBE RibbonCreator 2016                                #
'#      Version: 1.1002                                         #
'#                                                              #
'#      (c) 2007-2017 IDBE Avenius                              #
'#                                                              #
'#      http://www.ribboncreator2016.de/en                        #
'#      http://www.ribboncreator2010.com                      #
'#      http://www.ribboncreator.com                            #
'#      http://www.accessribon.com                              #
'#      http://www.avenius.com                                  #
'#                                                              #
'#      You may send change requests or report errors to:       #
'#      Aenderungswuensche oder Fehler bitte an:                #
'#                                                              #
'#      mailto://info@ribboncreator2016                         #
'#                                                              #
'################################################################


' Globals:


Global Const strAppPicturePath As String = "Pics"             ' The pictures/icons are available in the following directory
                                                              ' below the database directory
                                                              ' %Databasepath%\Pics
                                                              ' Die Bilder / Icons liegen in folgendem Verzeichnis
                                                              ' unterhalb des Datenbankverzeichnises
                                                              ' %Datenbankpfad%\Pics
                                                              
Global Const bolUseDynamicPicturePath As Boolean = False      ' The images should be loaded from this directory
                                                              ' and not from the directory in the Ribbon XML.
                                                              ' Die Bilder sollen aus diesem Verzeichnis geladen werden
                                                              ' und nicht ueber die Verzeichnisangabe im Ribbon XML.

                                                              ' These values are used in the function "GetImages"
                                                              ' Diese Werte werden in der Funktion "GetImages" verwendet.

                                                              
Public gobjRibbon As IRibbonUI

Public bolEnabled As Boolean    ' Used in Callback "getEnabled"
                                ' Further informations in Callback "getEnabled"
                                ' Fuer Callback "getEnabled"
                                ' Genauere Informationen in Callback "getEnabled".
                               
Public bolVisible As Boolean    ' Used in Callback "getVisible"
                                ' More information in Callback "getVisible
                                ' Fuer Callback "getVisible"
                                ' Further informations in Callback "getVisible

' For Sample Callback "GetContent"
' Fuer Beispiel Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type


' Callbacks:

Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' Callbackname in XML File "onLoad"

    Set gobjRibbon = ribbon
End Sub

Sub LoadImages(control, ByRef image)
    ' Callbackname in XML File "loadImage"

    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required
    ' Laed ein Bild mit Transparenz in das Ribbon
    ' Modul basGDIPlus wird dafuer benoetigt
    
    Dim strImage        As String
    Dim strPicture      As String
    
    strImage = CStr(control)
    strPicture = getPic(strImage)
    
    If strImage <> "" Then
        If bolUsePicturesFromTable = True Then
            If strPicture <> "" Then
                Set image = getIconFromTable(strPicture)
            Else
                Set image = Nothing
            End If
        Else
            Set image = LoadPictureGDIP(strImage)
        End If
    Else
        Set image = Nothing
    End If
    
End Sub

Sub GetImages(control As IRibbonControl, ByRef image)
    ' Callbackname in XML File "getImages"
    
    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required
    ' Laed ein Bild mit Transparenz in das Ribbon
    ' Modul basGDIPlus wird dafuer benoetigt
    
    Dim strPicturePath  As String
    Dim strPicture      As String
    
    strPicture = getTheValue(control.Tag, "CustomPicture")
    
    If bolUsePicturesFromTable = True Then
        Set image = getIconFromTable(strPicture)
    Else
        If bolUseDynamicPicturePath = True Then
            strPicturePath = getAppPath & strAppPicturePath & "\"
        Else
            strPicturePath = getTheValue(control.Tag, "CustomPicturePath")
        End If
        Set image = LoadPictureGDIP(strPicturePath & strPicture)
    End If

End Sub

Sub GetEnabled(control As IRibbonControl, ByRef enabled)
    ' Callbackname in XML File "getEnabled"
    
    ' To set the property "enabled" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Enabled Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        
        Case "btn_crnmtn"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            enabled = True
        Case "btn_crnset"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            enabled = True
        Case "btn_crndskexp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            enabled = True
        Case "btn_crndskimp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            enabled = True
        Case "mnu_crndsk"
            ' Menu
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            enabled = True
        Case "btn_crnstat"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            enabled = True
        Case "tgb_dev"
            ' ToggleButton
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            enabled = True
        Case "btn_setutil"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            enabled = True
        Case "btn_setteams"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            enabled = True
        Case "btn_setpoints"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            enabled = True
        Case "btn_compimp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            enabled = True
        Case "btn_compman"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            enabled = True
        Case "btn_evntdetail"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            enabled = True
        Case "btn_evntcomp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            enabled = True
        Case "btn_evntord"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            enabled = True
        Case "sep_14"
            ' Separator
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            enabled = True
        Case "btn_evntlist"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            enabled = True
        Case "btn_entres"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_entry
            enabled = True
        Case "btn_repstat"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_reports
            enabled = True
        Case "btn_repevntlist"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_reports
            enabled = True

        Case Else
            enabled = True

    End Select

End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    
    ' To set the property "visible" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Visible Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        
        Case "btn_crnmtn"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            visible = True
        Case "btn_crnset"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            visible = True
        Case "btn_crndskexp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            visible = True
        Case "btn_crndskimp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            visible = True
        Case "mnu_crndsk"
            ' Menu
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            visible = True
        Case "btn_crnstat"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            visible = True
        Case "tgb_dev"
            ' ToggleButton
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            visible = True
        Case "grp_crn"
            ' Group:    grp_crn
            ' In Tab:   tab_setup
            visible = True
        Case "btn_setutil"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            visible = True
        Case "btn_setteams"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            visible = True
        Case "btn_setpoints"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            visible = True
        Case "grp_setup"
            ' Group:    grp_setup
            ' In Tab:   tab_setup
            visible = True
        Case "btn_compimp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            visible = True
        Case "btn_compman"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            visible = True
        Case "grp_comp"
            ' Group:    grp_comp
            ' In Tab:   tab_setup
            visible = True
        Case "btn_evntdetail"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            visible = True
        Case "btn_evntcomp"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            visible = True
        Case "btn_evntord"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            visible = True
        Case "sep_14"
            ' Separator
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            visible = True
        Case "btn_evntlist"
            ' Button
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            visible = True
        Case "grp_evnt"
            ' Group:    grp_evnt
            ' In Tab:   tab_setup
            visible = True
        Case "tab_setup"
            ' Tab:   tab_setup
            visible = True
        Case "btn_entres"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_entry
            visible = True
        Case "grp_entry"
            ' Group:    grp_entry
            ' In Tab:   tab_results
            visible = True
        Case "btn_repstat"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_reports
            visible = True
        Case "btn_repevntlist"
            ' Button
            ' In Tab:   tab_results
            ' In Group: grp_reports
            visible = True
        Case "grp_reports"
            ' Group:    grp_reports
            ' In Tab:   tab_results
            visible = True
        Case "tab_results"
            ' Tab:   tab_results
            visible = True

        Case Else
            visible = True

    End Select

End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case control.id
        

        Case Else
            label = "*getLabel*"

    End Select

End Sub

Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case control.id
        

        Case Else
            screentip = "*getScreentip*"

    End Select

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case control.id
        

        Case Else
            supertip = "*getSupertip*"

    End Select

End Sub

Sub GetDescription(control As IRibbonControl, ByRef Description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "description" to a Ribbon Control

    Select Case control.id
        

        Case Else
            Description = "*getDescription*"

    End Select

End Sub

Sub GetTitle(control As IRibbonControl, ByRef Title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.id
        

        Case Else
            Title = "*getTitle*"

    End Select

End Sub

'Button

Sub OnActionButton(control As IRibbonControl)
    ' Callbackname in XML File "onAction"

    ' Callback for event button click
    ' Callback fuer Button Click
    
    Select Case control.id
        
        Case "btn_crnmtn"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
             DoCmd.OpenForm "Carnivals Maintain" 'Opens the customers form'
        Case "btn_crnset"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            DoCmd.OpenForm "Setup Carnival" 'Opens the customers form'
        Case "btn_crndskexp"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            DoCmd.OpenForm "ExportData" 'Opens the customers form'
        Case "btn_crndskimp"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            ' In Menu: mnu_crndsk
            DoCmd.OpenForm "Import Data" 'Opens the customers form'
        Case "btn_crnstat"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            DoCmd.OpenForm "Statiscal Reports" 'Opens the customers form'
        Case "btn_setutil"
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            DoCmd.OpenForm "Utilities" 'Opens the customers form'
        Case "btn_setteams"
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            DoCmd.OpenForm "House Summary" 'Opens the customers form'
        Case "btn_setpoints"
            ' In Tab:   tab_setup
            ' In Group: grp_setup
            DoCmd.OpenForm "PointScale" 'Opens the customers form'
        Case "btn_compimp"
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            DoCmd.OpenForm "Import Competitors" 'Opens the customers form'
        Case "btn_compman"
            ' In Tab:   tab_setup
            ' In Group: grp_comp
            DoCmd.OpenForm "CompetitorsSummary" 'Opens the customers form'
        Case "btn_evntdetail"
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            DoCmd.OpenForm "EventTypeSummary" 'Opens the customers form'
        Case "btn_evntcomp"
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            DoCmd.OpenForm "CompEventsSummary" 'Opens the customers form'
        Case "btn_evntord"
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            DoCmd.OpenForm "EventOrder" 'Opens the customers form'
        Case "btn_evntlist"
            ' In Tab:   tab_setup
            ' In Group: grp_evnt
            DoCmd.OpenForm "Reports_Event" 'Opens the customers form'
        Case "btn_entres"
            ' In Tab:   tab_results
            ' In Group: grp_entry
            DoCmd.OpenForm "CompEventsSummary" 'Opens the customers form'
        Case "btn_repstat"
            ' In Tab:   tab_results
            ' In Group: grp_reports
            DoCmd.OpenForm "Statiscal Reports" 'Opens the customers form'
        Case "btn_repevntlist"
            ' In Tab:   tab_results
            ' In Group: grp_reports
            DoCmd.OpenForm "Reports_Event" 'Opens the customers form'
      
        Case Else
            MsgBox "Button """ & control.id & """ clicked", vbInformation
    End Select

End Sub

'Command Button

Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
    ' Callbackname in XML File Command "onAction"

    ' Callback for command event button click
    ' Callback fuer Command Button Click

    MsgBox "Button ""Help"" clicked" & vbCrLf, _
                           vbInformation
    CancelDefault = True

End Sub

'CheckBox

Sub OnActionCheckBox(control As IRibbonControl, _
                     pressed As Boolean)
    ' Callbackname in XML File "OnActionCheckBox"
    
    ' Callback for event checkbox click
    ' Callback fuer Checkbox Click

    Select Case control.id
        
        
        Case Else
            MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf, _
                   vbInformation

    End Select

End Sub

Sub GetPressedCheckBox(control As IRibbonControl, _
                       ByRef bolReturn)
    ' Callbackname in XML File "GetPressedCheckBox"
    
    ' Callback for checkbox
    ' indicates how the control is displayed
    ' Callback fuer Checkbox wie das Control
    ' angezeigt werden soll

    Select Case control.id
        

        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If

    End Select

End Sub

'ToggleButton

Sub OnActionTglButton(control As IRibbonControl, _
                      pressed As Boolean)
                              
    ' Callbackname in XML File "onAction"
    
    ' Callback fuer einen Toggle Button Klick
    ' Callback for a Toggle Buttons click event

    Select Case control.id
        
        Case "tgb_dev"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed, _
                   vbInformation

        Case Else
            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed, _
                   vbInformation

    End Select

End Sub

Sub GetPressedTglButton(control As IRibbonControl, _
                        ByRef pressed)
    ' Callbackname in XML File "getPressed"

    ' Callback for an Access ToogleButton Control. Indicates how the control is displayed

    Select Case control.id
        
        Case "tgb_dev"
            ' In Tab:   tab_setup
            ' In Group: grp_crn
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If

        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If

    End Select

End Sub

'EditBox

Sub GetTextEditBox(control As IRibbonControl, _
                   ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    
    ' Callback fuer EditBox welcher Wert in der
    ' EditBox eingetragen werden soll.
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case control.id
        

        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    
    End Select
    
End Sub

Sub OnChangeEditBox(control As IRibbonControl, _
                    strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    
    ' Callback Editbox: Rueckgabewert der Editbox
    ' Callback Editbox: Return value of the Editbox

    Select Case control.id
        

        Case Else
            MsgBox "The Value of the EditBox """ & control.id & """ is: " & strText & vbCrLf & _
                   "Der Wert der EditBox """ & control.id & """ ist: " & strText, _
                   vbInformation

    End Select

End Sub

'DropDown

Sub OnActionDropDown(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    
    ' Callback onAction (DropDown)
    
    Select Case control.id
        

        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des DropDown-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub


Sub GetSelectedItemIndexDropDown(control As IRibbonControl, _
                                 ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    
    ' Callback getSelectedItemIndex
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            

            Case Else
                index = getTheValue(control.Tag, "DefaultValue")

        End Select

    End If

End Sub

'Gallery

Sub GetSelectedItemIndexGallery(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    
    ' Callback GetSelectedItemIndexGallery
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            

            Case Else
                index = varIndex

        End Select

    End If

End Sub

Sub OnActionGallery(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    
    ' Callback onAction (Gallery)
    
    Select Case control.id
        

        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des Gallery-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

'Combobox

Sub GetTextComboBox(control As IRibbonControl, _
                      ByRef strText)

    ' Callbackname im XML File "GetTextComboBox"
    
    ' Callback getText (Combobox)
                           
    Select Case control.id
        

        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select

End Sub


Sub OnChangeComboBox(control As IRibbonControl, _
                               strText As String)
                           
    ' Callbackname im XML File "OnChangeCombobox"
    
    ' Callback onChange (Combobox)
   
    Select Case control.id
        

        Case Else
            MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & strText & """" & vbCrLf & _
                   "Das selektierte Item des Combobox-Control """ & control.id & """ ist : """ & strText & """", _
                   vbInformation
    End Select

End Sub

' DynamicMenu

Sub GetContent(control As IRibbonControl, _
               ByRef XMLString)

    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    ' Beispiel fuer einen Ribbon XML - "getContent" Callback
    ' Siehe auch: http://www.accessribbon.de/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '       und : http://www.accessribbon.de/?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case control.id
        

        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
 
End Sub


' Helper Function
' Hilfsfunktionen

Public Function getXMLForDynamicMenu() As String
    
    ' Creates a XML String for DynamicMenu CallBack - getContent
    
    ' Erstellt den Inhalt fuer das DynamicMenu im Callback getContent
    
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next
 

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy

End Function

Public Function getTheValue(strTag As String, strValue As String) As String
    ' *************************************************************
    ' Created from     : Avenius
    ' Parameter        : Input String, SuchValue String
    ' Date created     : 05.01.2008
    '
    ' Sample:
    ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
    ' Return           : "Test"
    ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function

Public Function getAppPath() As String
    Dim strDummy As String
    strDummy = CurrentProject.Path
    If Right(strDummy, 1) <> "\" Then strDummy = strDummy & "\"
    getAppPath = strDummy
End Function

Public Function getIconFromTable(strFileName As String) As Picture
'*****************************************************************************
'Funktion 'getIconFromTable' holt ein Bild aus der Binaer-Tabelle und gibt
' ein Picture Objekt zurueck
' strFilename ist Bildobjekt welches in der Datei 'tblBinary" vorhanden ist.
'*****************************************************************************

Dim lSize As Long
Dim arrBin() As Byte
Dim rs As DAO.Recordset
 
    On Error GoTo Errr
 
    Set rs = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    rs.FindFirst "[FileName]='" & strFileName & "'"
    If rs.NoMatch Then
        Set getIconFromTable = Nothing
    Else
        lSize = rs.Fields("binary").FieldSize
        ReDim arrBin(lSize)
        arrBin = rs.Fields("binary").GetChunk(0, lSize)
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    rs.Close
 
fExit:
    Reset
    Erase arrBin
    Set rs = Nothing
    Exit Function
Errr:
    Resume fExit
End Function

Public Function getPic(strFullPath As String) As String
    Dim strResult As String
    
    If InStrRev(strFullPath, "\") > 0 Then
        strResult = Mid(strFullPath, InStrRev(strFullPath, "\") + 1)
    Else
        strResult = ""
    End If
   
    getPic = strResult
End Function