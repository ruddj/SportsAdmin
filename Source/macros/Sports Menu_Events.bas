﻿Version =0
ColumnsShown =1
Begin
    MacroName ="Maintain Event &Details"
    Action ="RunCode"
    Argument ="OF(\"EventTypeSummary\",\"M\")"
End
Begin
    MacroName ="Maintain &Competitors in Events"
    Action ="RunCode"
    Argument ="OF(\"CompEventsSummary\",\"M\")"
End
Begin
    MacroName ="Maintain Event &Order"
    Action ="RunCode"
    Argument ="OF(\"EventOrder\",\"M\")"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="&Generate Event Lists"
    Action ="RunCode"
    Argument ="OF(\"Reports_Event\",\"N\")"
End