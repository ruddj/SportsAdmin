﻿Version =0
ColumnsShown =1
Begin
    MacroName ="&Print"
    Action ="PrintOut"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="0"
    Argument ="1"
    Argument ="-1"
End
Begin
    MacroName ="Print Pre&view"
    Condition ="CurrentUser()=\"Owner\""
    Action ="DoMenuItem"
    Argument ="20"
    Argument ="1"
    Argument ="0"
    Argument ="11"
    Argument ="0"
End
Begin
    MacroName ="Print &Setup"
    Condition ="CurrentUser()=\"Owner\""
    Action ="DoMenuItem"
    Argument ="20"
    Argument ="1"
    Argument ="0"
    Argument ="10"
    Argument ="0"
End
