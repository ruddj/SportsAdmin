﻿Version =0
ColumnsShown =3
Begin
    MacroName ="&Maintain Carnivals"
    Action ="RunCode"
    Argument ="OF(\"Carnivals Maintain\",\"M\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="&Setup Carnival"
    Action ="RunCode"
    Comment ="Setup a sports carnival"
    Argument ="OF(\"Setup Carnival\",\"N\")"
End
Begin
    MacroName ="&Create Carnival Disks"
    Action ="RunCode"
    Argument ="OF(\"ExportData\",\"M\")"
End
Begin
    MacroName ="&Import Carnival Disks"
    Action ="RunCode"
    Argument ="OF(\"Import Data\",\"M\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="Maintain &Houses"
    Action ="RunCode"
    Argument ="OF(\"House Summary\",\"M\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="Maintain &Pointscales"
    Action ="RunCode"
    Argument ="OF(\"PointScale\",\"M\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="Carnival S&tatistics"
    Action ="RunCode"
    Argument ="OF(\"Statiscal Reports\",\"N\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="&Quit"
    Action ="Quit"
    Argument ="1"
End
Begin
    MacroName ="-"
    Condition ="UCase(CurrentUser())=\"OWNER\""
End
Begin
    MacroName ="&Restore"
    Condition ="UCase(CurrentUser())=\"OWNER\""
    Action ="DoMenuItem"
    Argument ="20"
    Argument ="1"
    Argument ="4"
    Argument ="4"
    Argument ="0"
End
Begin
    Condition ="UCase(CurrentUser())=\"OWNER\""
    Action ="SetValue"
    Argument ="[Application].[MenuBar]"
    Argument ="\"\""
End
