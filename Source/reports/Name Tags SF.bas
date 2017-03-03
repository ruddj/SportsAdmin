Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridX =25
    GridY =25
    Width =5669
    ItemSuffix =11
    Left =2265
    Top =990
    RecSrcDt = Begin
        0xd68ba1ac52c7e140
    End
    RecordSource ="SELECT DISTINCTROW CompEvents.PIN, EventType.ET_Des, CompEvents.Lane, CompEvents"
        ".Heat, CompEvents.F_Lev, Heats.E_Number FROM EventType INNER JOIN (Heats INNER J"
        "OIN (Events INNER JOIN CompEvents ON Events.E_Code = CompEvents.E_Code) ON (Heat"
        "s.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev) AND (Heats.E_Code"
        " = CompEvents.E_Code)) ON EventType.ET_Code = Events.ET_Code WHERE (((CompEvents"
        ".F_Lev)=0) AND ((EventType.Flag)=True)) ORDER BY CompEvents.PIN;"
    OnOpen ="[Event Procedure]"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000251600005a00000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="PIN"
        End
        Begin BreakLevel
            ControlSource ="ET_Des"
        End
        Begin PageHeader
            Height =226
            Name ="PageHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader3"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =90
            Name ="Detail1"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    Left =735
                    Width =2841
                    Height =60
                    FontSize =18
                    Name ="Name"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =696
                    Height =60
                    FontSize =18
                    TabIndex =1
                    BackColor =12632256
                    Name ="Number"
                    ControlSource ="E_Number"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter2"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nFontSize As Variant

Private Sub Report_Open(Cancel As Integer)

  nFontSize = Val(DLookup("[NameTagFontSize]", "Misc-EventLists"))
  Me!Name.FontSize = nFontSize
  Me!Number.FontSize = nFontSize
  
End Sub
