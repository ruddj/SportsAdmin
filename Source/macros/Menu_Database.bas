﻿Version =0
ColumnsShown =3
Begin
    MacroName ="&Quit"
    Action ="Quit"
    Argument ="1"
End
Begin
    MacroName ="&Restore"
    Condition ="CurrentUser()=\"Owner\""
    Action ="DoMenuItem"
    Argument ="20"
    Argument ="1"
    Argument ="4"
    Argument ="4"
    Argument ="0"
End
Begin
    Condition ="CurrentUser()=\"Owner\""
    Action ="SetValue"
    Argument ="[Application].[MenuBar]"
    Argument ="\"\""
End
