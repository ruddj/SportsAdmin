﻿Version =0
ColumnsShown =1
Begin
    MacroName ="&Setup"
    Action ="OpenForm"
    Argument ="Setup Carnival"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="&Create Carnival Disks"
    Condition ="CurrentUser()=\"Owner\""
    Action ="OpenForm"
    Argument ="ExportData"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="&Import Carnival Disks"
    Condition ="CurrentUser()=\"Owner\""
    Action ="OpenForm"
    Argument ="Import Data"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="&Maintain Carnivals"
    Action ="OpenForm"
    Argument ="Carnivals Maintain"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
