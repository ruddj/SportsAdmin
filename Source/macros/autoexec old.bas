﻿Version =0
ColumnsShown =2
Begin
    Action ="SetValue"
    Argument ="[Application].[MenuBar]"
    Argument ="\"Sports Menu\""
End
Begin
    Action ="RunCommand"
    Argument ="2"
End
Begin
    Action ="RunCode"
    Argument ="InitialiseWaitMessage () "
End
Begin
    Action ="RunMacro"
    Argument ="ShowPleaseWait"
End
Begin
    Action ="ShowToolbar"
    Argument ="Database"
    Argument ="2"
End
Begin
    Action ="ShowToolbar"
    Argument ="Form View"
    Argument ="2"
End
Begin
    Action ="ShowToolbar"
    Argument ="Print Preview"
    Argument ="1"
End
Begin
    Action ="RunCode"
    Argument ="CheckInventoryAttached()"
End
Begin
    Action ="RunCode"
    Argument ="Startup()"
End
Begin
    Action ="RunMacro"
    Argument ="ClosePleaseWait"
End
Begin
    Action ="OpenForm"
    Argument ="Main Menu"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End