Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10712
    ItemSuffix =116
    Left =4470
    Top =825
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xea037adf7725e240
    End
    RecordSource ="House points - GrandTotal"
    Caption ="Overall Statistics"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x350200003502000035020000d002000000000000d8290000d401000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    RibbonName ="SportPrint"
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="GrandTotal"
        End
        Begin BreakLevel
            ControlSource ="H_NAme"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader0"
        End
        Begin PageHeader
            Height =1787
            OnFormat ="[Event Procedure]"
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10656
                    Height =450
                    FontSize =16
                    FontWeight =700
                    Name ="CarnivalTitle"
                    ControlSource ="=DLookUp(\"[CarnivalTitle]\",\"Miscellaneous\")"
                    FontName ="times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =3
                    BorderLineStyle =3
                    Left =56
                    Top =566
                    Width =10596
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Left =120
                    Top =735
                    Width =8355
                    Height =390
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="Overall Statistical Summary - Ordered by Grand Total"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =850
                    Top =1370
                    Width =1365
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text107"
                    Caption ="Name"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =4195
                    Top =1360
                    Width =1020
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text108"
                    Caption ="Team"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =5499
                    Top =1360
                    Width =960
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text109"
                    Caption ="Total"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6402
                    Top =1360
                    Width =1080
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text110"
                    Caption ="Extras"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =7766
                    Top =1360
                    Width =1170
                    Height =360
                    FontSize =11
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Grand Total"
                    FontName ="Times New Roman"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =56
                    Top =1757
                    Width =10647
                    Name ="Line112"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =9238
                    Top =1365
                    Width =1245
                    Height =330
                    FontSize =11
                    FontWeight =700
                    Name ="Text114"
                    Caption ="% Total"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =468
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    Left =56
                    Top =453
                    Width =10647
                    Name ="Line86"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =850
                    Top =56
                    Width =3246
                    Height =330
                    FontSize =11
                    FontWeight =700
                    Name ="Field98"
                    ControlSource ="H_NAme"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4078
                    Top =56
                    Width =1131
                    Height =330
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    Name ="Field100"
                    ControlSource ="H_Code"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5272
                    Top =56
                    Width =921
                    Height =330
                    FontSize =11
                    TabIndex =2
                    Name ="SumOfPoints"
                    ControlSource ="SumOfPoints"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =6179
                    Top =56
                    Width =1026
                    Height =330
                    FontSize =11
                    TabIndex =3
                    Name ="Extras"
                    ControlSource ="Extras"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =7370
                    Top =56
                    Width =1236
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="GrandTotal"
                    ControlSource ="GrandTotal"

                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =170
                    Top =56
                    Width =636
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="Place"
                    ControlSource ="=1"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontFamily =34
                    Left =8894
                    Top =56
                    Width =1356
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="PercentileTotal"
                    ControlSource ="PercentileTotal"
                    Format ="Percent"

                End
            End
        End
        Begin PageFooter
            Height =6059
            OnFormat ="[Event Procedure]"
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9504
                    Top =5616
                    Width =1131
                    Height =435
                    FontSize =6
                    TabIndex =2
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =5669
                    Width =9651
                    Height =390
                    FontSize =11
                    FontWeight =700
                    Name ="Field88"
                    ControlSource ="=DLookUp(\"[CarnivalFooter]\",\"Miscellaneous\")"
                    FontName ="Times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =2
                    BorderLineStyle =3
                    Top =5669
                    Width =10596
                    Name ="Line87"
                End
                Begin Chart
                    ColumnHeads = NotDefault
                    SizeMode =3
                    RowSourceTypeInt =2
                    Left =1133
                    Top =113
                    Width =8677
                    Height =5371
                    TabIndex =1
                    Name ="oleChart"
                    OleData = Begin
                        0x00440000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff050000000108020000000000c0000000 ,
                        0x00000046000000000000000000000000000000000000000007000000400b0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000063000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002c00000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000d7080000 ,
                        0x0000000005000000fdfffffffffffffffffffffffffffffffefffffffeffffff ,
                        0x12000000ffffffffffffffff150000000c0000000d0000000e00000018000000 ,
                        0xffffffffffffffffffffffff13000000140000000a000000feffffffffffffff ,
                        0xffffffff190000001a0000001b0000001f000000ffffffffffffffffffffffff ,
                        0x20000000feffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff050000000308020000000000c0000000 ,
                        0x0000004600000000000000000000000000aa799285afbd0109000000400b0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000068000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002c00000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000d7080000 ,
                        0x00000000ffffffffffffffff05000000fdfffffffefffffffeffffffffffffff ,
                        0xfffffffffeffffff12000000080000000c0000000d0000000e00000018000000 ,
                        0xffffffffffffffffffffffff13000000140000000a000000ffffffffffffffff ,
                        0xffffffff190000001a0000001b0000001f000000ffffffffffffffffffffffff ,
                        0x20000000feffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x20000000210000002200000023000000240000002500000026000000feffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffeffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff03004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000180002010300000004000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000b000000dc120000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x20000000210000002200000023000000240000002500000026000000feffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffeffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0108020000000000c000000000000046 ,
                        0x140000004d6963726f736f667420477261706820352e30000700000047426966 ,
                        0x663500100000004d5347726170682e43686172742e3500f439b2710000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000908080080050500e209c9071e041a00050017222422232c2323305f ,
                        0x293b5c28222422232c2323305c291e041f0006001c222422232c2323305f293b ,
                        0x5b5265645d5c28222422232c2323305c291e04200007001d222422232c232330 ,
                        0x2e30305f293b5c28222422232c2323302e30305c291e04250008002222242223 ,
                        0x2c2323302e30305f293b5b5265645d5c28222422232c2323302e30305c291e04 ,
                        0x35002a00325f282224222a20232c2323305f293b5f282224222a205c28232c23 ,
                        0x23305c293b5f282224222a20222d225f293b5f28405f291e042c002900295f28 ,
                        0x2a20232c2323305f293b5f282a205c28232c2323305c293b5f282a20222d225f ,
                        0x293b5f28405f291e043d002c003a5f282224222a20232c2323302e30305f293b ,
                        0x5f282224222a205c28232c2323302e30305c293b5f282224222a20222d223f3f ,
                        0x5f293b5f32100400000003003310000007100a00000000000000000009000a10 ,
                        0x0c00ffffff000000000000000000341000003410000024100200020025101a00 ,
                        0x0202010000000000ddffffffc7ffffff0000000000000000b100331000004f10 ,
                        0x1400020002000000000000000000000000000000000026100200020051100800 ,
                        0x00010200ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xbd3b00001b250000540f0000010009000003aa07000007001c00000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c0267014202030000001e00040000002701ffff050000000b02 ,
                        0x0000000001000002040000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000001000 ,
                        0x0000fb02f5ff000000000000bc020000000000000022417269616c00a8810400 ,
                        0x00002d01000010000000fb021000070000000000bc0200000000010202225379 ,
                        0x7374656d006e040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016046701 ,
                        0x420200000000040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d01000038000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0308020000000000c000000000000046 ,
                        0x190000004d6963726f736f667420477261706820393720436861727400070000 ,
                        0x0047426966663500100000004d5347726170682e43686172742e3800f439b271 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000908080080050500e209c9071e041a00050017222422232c2323305f ,
                        0x293b5c28222422232c2323305c291e041f0006001c222422232c2323305f293b ,
                        0x5b5265645d5c28222422232c2323305c291e04200007001d222422232c232330 ,
                        0x2e30305f293b5c28222422232c2323302e30305c291e04250008002222242223 ,
                        0x2c2323302e30305f293b5b5265645d5c28222422232c2323302e30305c291e04 ,
                        0x35002a00325f282224222a20232c2323305f293b5f282224222a205c28232c23 ,
                        0x23305c293b5f282224222a20222d225f293b5f28405f291e042c002900295f28 ,
                        0x2a20232c2323305f293b5f282a205c28232c2323305c293b5f282a20222d225f ,
                        0x293b5f28405f291e043d002c003a5f282224222a20232c2323302e30305f293b ,
                        0x5f282224222a205c28232c2323302e30305c293b5f282224222a20222d223f3f ,
                        0x5f293b5f00000000000000000f00151014002b0e0000ab00000075010000b201 ,
                        0x000007011200331000004f101400050002002b0e0000ac000000000000000000 ,
                        0x000025101a000202010000000000ddffffffc7ffffff0000000000000000b100 ,
                        0x331000004f101400020002000000000000000000000000000000000051100800 ,
                        0x00010200000000003410000032100400000003003310000007100a0000000000 ,
                        0x0000000009000a100c00ffffff00000000000000000034100000341000002410 ,
                        0x0200020025101a000202010000000000ddffffffc7ffffff0000000000000000 ,
                        0xb100331000004f10140002000200000000000000000000000000000000002610 ,
                        0x0200020051100800000102000000000034100000341000003410000025101a00 ,
                        0x020201000000000093060000500000007a0200001e0100008100331000004f10 ,
                        0x14000200020000000000000000005a0000001900000051100800000102000000 ,
                        0x00000d10110000000e4f766572616c6c2053636f726573271006000100000000 ,
                        0x00341000003410000000000a00000008000000020000000a0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000ddffffffc7ffffff0000000000000000b100331000004f101400 ,
                        0x0200020000000000000000000000000000000000511008000001020000000000 ,
                        0x34100000ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xbd3b000035250000b41200000100090000035a09000008000502000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c0268014202030000001e0005020000f7000003000100000000 ,
                        0x800000000080000080800000000080008000800000808000c0c0c000c0dcc000 ,
                        0xa6caf00080000004008000048080000400008004800080040080800480808004 ,
                        0x8080ff0480206004ffffc004a0e0e00460008004ff8080040080c004c0c0ff04 ,
                        0x00cfff0469ffff04e0ffe004dd9cb304b38fee042a6ff9043fb8cd0448843604 ,
                        0x958c41048e5e4204a0627a04624fac041d2fbe042866760400450004453e0104 ,
                        0x6a28130485396a044a32850404040404080808040c0c0c041111110416161604 ,
                        0x1c1c1c042222220429292904303030045f5f5f04555555044d4d4d0442424204 ,
                        0x39393904000700040d000004b79981048499b404bdbd90047f7f600460607f04 ,
                        0x000e00041b0000042800000408092b04001d00043900000400009b0400250004 ,
                        0x4900000411113b04002f00045d00000417174504003a0004491111041c1c5304 ,
                        0x0016ff042b00ff0421216c045914140400510004471a6a041932670400610004 ,
                        0x0031ff046100ff0453207b04164367042e2ee2042659160451460404682e4904 ,
                        0x07528f046a18b804902315040053ff04a300ff046a4a120475336c044a419a04 ,
                        0x37650b04a42c1504831fb1044e2cff042051b604086492046f560b045943ad04 ,
                        0x36721204b033170400a10004775f1f0489477104b0431c04b72d7d0400869504 ,
                        0x7a6e2304269f000473a901040000000400000004000000040000000400000004 ,
                        0x00000004000000040000000400ca0004ac5b0104201dc2049452700424aa4c04 ,
                        0x0a948904366e7b0444759004ff00a8040071ff04df00ff0456914a043448f804 ,
                        0xcc328204e441700468ca010436bc4204009aff049622b704857d330425b78c04 ,
                        0x365aed045cff0004ff480004229ba20442cf4d04c258520420d39504a524e004 ,
                        0x7356b504a9a90004d06f3c04679f580489cf0b04ffac0004a72efe04e2597f04 ,
                        0x4cdc6704ff18ff043a7dff04b1d01804c7ff0004ffe20004df9a3d0456819f04 ,
                        0xc643ba04af718b0438a2c904d153ce04ff9a650446cadb04ff4dff04c8e96a04 ,
                        0x4cdee004ff98ff04dfc08204e9eca504f5f6cd04ffd0ff04b1ac5a046391ae04 ,
                        0x224c65048d4e3f0450707004d0ffff04ffe7ff04696969047777770486868604 ,
                        0x969696049d9d9d04a4a4a404b2b2b204cbcbcb04d7d7d704dddddd04e3e3e304 ,
                        0xeaeaea04f1f1f104f8f8f804b2c1660480bf7804c6f0f004b2a4ff04ffb3ff04 ,
                        0xd18ea304c3dc3704a09e540476ae7004789ec1048364bf04a483d304d13f3204 ,
                        0xff7d000444782304245f60040e0e2c04be000004ff1f000431390004d9853e04 ,
                        0x02778504b0d8810456211d040000300488c8b304a0790004c0c0c004ea708104 ,
                        0x51f16904ffff80049174cd04ff7cff04a2ffff04fffbf000a0a0a40080808000 ,
                        0xff00000000ff0000ffff00000000ff00ff00ff0000ffff00ffffff0004000000 ,
                        0x34020000030000003500040000002701ffff050000000b020000000003000000 ,
                        0x1e00050000000102ffffff000500000009020000000010000000fb02f5ff0000 ,
                        0x00000000bc020000000000000022417269616c00c100040000002d0101001000 ,
                        0x0000fb021000070000000000bc02000000000102022253797374656d006e0400 ,
                        0x00002d010200040000002701ffff030000001e00040000002d01010005000000 ,
                        0x0102ffffff000500000009020000000007000000160468014202000000000400 ,
                        0x00002d010200040000002701ffff030000001e00040000002d01010005000000 ,
                        0x0102ffffff000500000009020000000007000000160468014202000000000700 ,
                        0x000015043a003d0214000802040000002d010200040000002701ffff03000000 ,
                        0x1e00040000002d010100050000000102ffffff00050000000902000000000700 ,
                        0x000016046801420200000000040000002d010200040000002701ffff03000000 ,
                        0x1e00040000002d010100050000000102ffffff00050000000902000000000700 ,
                        0x0000160464013e0204000400040000002d010200040000002701ffff03000000 ,
                        0x1e00040000002d010100050000000102ffffff00050000000902000000000700 ,
                        0x00001604400112023f002100040000002d010200040000002701ffff03000000 ,
                        0x1e00040000002d010100050000000102ffffff00050000000902000000000700 ,
                        0x000016043f0113023d00210009000000fa020000010000000000000022000400 ,
                        0x00002d01030007000000fc0200008080ff02ffff040000002d01040005000000 ,
                        0x0902ffffff000500000001028080ff020400000004010d000400000002010200 ,
                        0x0e00000024030500470010017900100179003e0147003e01470010010e000000 ,
                        0x24030500c300fe00f400fe00f4003e01c3003e01c300fe000e00000024030500 ,
                        0x3e01f1007001f10070013e013e013e013e01f1000e00000024030500ba01e400 ,
                        0xeb01e400eb013e01ba013e01ba01e40007000000fc02000080206002ffff0400 ,
                        0x00002d01050004000000f0010400050000000102802060020e00000024030500 ,
                        0x3e014800700148007001f1003e01f1003e01480009000000fa02000000000000 ,
                        0x000000002200040000002d01040007000000fc020000ffffff00000004000000 ,
                        0x2d010600040000002d010200040000002701ffff030000001e00040000002d01 ,
                        0x0100040000002d010500040000002d0103000500000001028020600205000000 ,
                        0x0902ffff00000000000000000f00151014002b0e0000ab00000075010000b201 ,
                        0x000007011200331000004f101400050002002b0e0000ac000000000000000000 ,
                        0x000025101a000202010000000000ddffffffc7ffffff0000000000000000b100 ,
                        0x331000004f101400020002000000000000000000000000000000000051100800 ,
                        0x00010200000000003410000032100400000003003310000007100a0000000000 ,
                        0x0000000009000a100c00ffffff00000000000000000034100000341000002410 ,
                        0x0200020025101a000202010000000000ddffffffc7ffffff0000000000000000 ,
                        0xb100331000004f10140002000200000000000000000000000000000000002610 ,
                        0x0200020051100800000102000000000034100000341000003410000025101a00 ,
                        0x020201000000000093060000500000007a0200001e0100008100331000004f10 ,
                        0x14000200020000000000000000005a0000001900000051100800000102000000 ,
                        0x00000d10110000000e4f766572616c6c2053636f726573271006000100000000 ,
                        0x00341000003410000000000a00000008000000020000000a0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000ddffffffc7ffffff0000000000000000b100331000004f101400 ,
                        0x0200020000000000000000000000000000000000511008000001020000000000 ,
                        0x3410000028405f291e0434002b00315f282a20232c2323302e30305f293b5f28 ,
                        0x2a205c28232c2323302e30305c293b5f282a20222d223f3f5f293b5f28405f29 ,
                        0x1e041600a40013222422232c2323303b5c2d222422232c2323301e041b00a500 ,
                        0x18222422232c2323303b5b5265645d5c2d222422232c2323301e041c00a60019 ,
                        0x222422232c2323302e30303b5c2d222422232c2323302e30301e042100a7001e ,
                        0x222422232c2323302e30303b5b5265645d5c2d222422232c2323302e30301e04 ,
                        0x3300a800305f2d2224222a20232c2323305f2d3b5c2d2224222a20232c232330 ,
                        0x5f2d3b5f2d2224222a20222d225f2d3b5f2d405f2d1e042a00a900275f2d2a20 ,
                        0x232c2323305f2d3b5c2d2a20232c2323305f2d3b5f2d2a20222d225f2d3b5f2d ,
                        0x405f2d1e043b00aa00385f2d2224222a20232c2323302e30305f2d3b5c2d2224 ,
                        0x222a20232c2323302e30305f2d3b5f2d2224222a20222d223f3f5f2d3b5f2d40 ,
                        0x5f2d1e043200ab002f5f2d2a20232c2323302e30305f2d3b5c2d2a20232c2323 ,
                        0x302e30305f2d3b5f2d2a20222d223f3f5f2d3b5f2d405f2d31001400a0000100 ,
                        0xff7fbc0200000000001205417269616c31001400c8000100ff7fbc0200000000 ,
                        0x001205417269616c31001400a0000100ff7fbc0200000002001205417269616c ,
                        0x31001400c8000100ff7fbc0200000000001205417269616c3d001200a005c003 ,
                        0x5622ad160000264d616a6f72004e85000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x0001002610020003005310040000000500541004000000030055100600000000 ,
                        0x00000104000e000000000000000006485f436f646504000e0000000100000000 ,
                        0x06506f696e747304000e00000002000000000645787472617304000d00010000 ,
                        0x00000000054a5544414803000f0001000100000000000000000080544003000f ,
                        0x0001000200000000000000000000000004000d00020000000000000541534845 ,
                        0x5203000f00020001000000000000000000805c4003000f000200020000000000 ,
                        0x0000000000000004000f0003000000000000074550485241494d03000f000300 ,
                        0x0100000000000000000000614003000f00030002000000000000000000c07240 ,
                        0x04000c0004000000000000044c45564903000f00040001000000000000000000 ,
                        0x00644003000f00040002000000000000000000000000571001000159100800b4 ,
                        0x002c01e62858203d000a00000000003b1fb80b00003e000e0001010100010100 ,
                        0x010001000000005810020000001d001100030100010000000100010001000100 ,
                        0x010034100000011002000000021010000000000000000000e87fb10100000e01 ,
                        0x33100000a0000400010001000310080003000100040004003310000051100800 ,
                        0x0001020028405f291e0434002b00315f282a20232c2323302e30305f293b5f28 ,
                        0x2a205c28232c2323302e30305c293b5f282a20222d223f3f5f293b5f28405f29 ,
                        0x1e041600a40013222422232c2323303b5c2d222422232c2323301e041b00a500 ,
                        0x18222422232c2323303b5b5265645d5c2d222422232c2323301e041c00a60019 ,
                        0x222422232c2323302e30303b5c2d222422232c2323302e30301e042100a7001e ,
                        0x222422232c2323302e30303b5b5265645d5c2d222422232c2323302e30301e04 ,
                        0x3300a800305f2d2224222a20232c2323305f2d3b5c2d2224222a20232c232330 ,
                        0x5f2d3b5f2d2224222a20222d225f2d3b5f2d405f2d1e042a00a900275f2d2a20 ,
                        0x232c2323305f2d3b5c2d2a20232c2323305f2d3b5f2d2a20222d225f2d3b5f2d ,
                        0x405f2d1e043b00aa00385f2d2224222a20232c2323302e30305f2d3b5c2d2224 ,
                        0x222a20232c2323302e30305f2d3b5f2d2224222a20222d223f3f5f2d3b5f2d40 ,
                        0x5f2d1e043200ab002f5f2d2a20232c2323302e30305f2d3b5c2d2a20232c2323 ,
                        0x302e30305f2d3b5f2d2a20222d223f3f5f2d3b5f2d405f2d31001400a0000100 ,
                        0xff7fbc0200000000006405417269616c31001400c8000100ff7fbc0200000000 ,
                        0x006405417269616c31001400a0000100ff7fbc0200000002006405417269616c ,
                        0x31001400c8000100ff7fbc0200000000006405417269616c3d001200a005c003 ,
                        0x5622ad160000cefb00008288020085000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x0001002610020003005310040000000500541004000000030055100600000000 ,
                        0x00000104000e000000000000000006485f436f646504000e0000000100000000 ,
                        0x06506f696e747304000e00000002000000000645787472617304000d00010000 ,
                        0x0000000005415348455203000f0001000100000000000000000070754003000f ,
                        0x0001000200000000000000000000000004000f00020000000000000745504852 ,
                        0x41494d03000f0002000100000000000000000070764003000f00020002000000 ,
                        0x00000000000000000004000d0003000000000000054a5544414803000f000300 ,
                        0x01000000000000000000e0774003000f00030002000000000000000000000000 ,
                        0x04000c0004000000000000044c45564903000f00040001000000000000000000 ,
                        0x50784003000f00040002000000000000000000000000571001000159100800b4 ,
                        0x002c01e62858203d000a00000000003b1fb80b00003e000e0001010100010100 ,
                        0x010001000000005810020000001d001100030100010000000100010001000100 ,
                        0x010034100000011002000000021010000000000000000000e87fb10100000e01 ,
                        0x33100000a0000400010001000310080003000100040004003310000051100800 ,
                        0x00010200000001000d100900000006506f696e74735110080001010200000001 ,
                        0x0051100800020102000000000006100800ffff00000000000045100200000034 ,
                        0x100000031008000300010004000400331000005110080000010200000002000d ,
                        0x1009000000064578747261735110080001010200000002005110080002010200 ,
                        0x0000000006100800ffff01000100000045100200000034100000031008000300 ,
                        0x010008000800331000005110080000010200000003000d100d0000000a477261 ,
                        0x6e64546f74616c51100800010102000000030051100800020102000000000006 ,
                        0x100800ffff020002000000451002000000341000004410030009000046100200 ,
                        0x0100411012000000cc000000a2020000970d0000570b0000331000004f101400 ,
                        0x02000200060000002c020000600e0000030d00001d1012000000000000000000 ,
                        0x00000000000000000000331000001e101a000200030100000000000000000000 ,
                        0x000000000000000000002300341000001d101200010000000000000000000000 ,
                        0x000000000000331000001f102a00000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000001e011e101a0002000301 ,
                        0x0000000000000000000000000000000000000000230034100000141014000000 ,
                        0x00000000000000000000000000000000000033100000171006009cff96000200 ,
                        0x22100a0032100400000003003310000007100a00000000000000000009000a10 ,
                        0x0c00ffffff000000000000000000341000003410000024100200020025101a00 ,
                        0x0202010000000000ddffffffc7ffffff0000000000000000b100331000004f10 ,
                        0x1400020002000000000000000000000000000000000026100200020051100800 ,
                        0x00010200ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xbd3b00001b250000540f0000010009000003aa07000007001c00000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c0267014202030000001e00040000002701ffff050000000b02 ,
                        0x0000000001000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000001000 ,
                        0x0000fb02f5ff000000000000bc020000000000000022417269616c00a8810400 ,
                        0x00002d01000010000000fb021000070000000000bc0200000000010202225379 ,
                        0x7374656d006e040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016046701 ,
                        0x420200000000040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000000001000d100900000006506f696e74735110080001010200000001 ,
                        0x0051100800020102000000000006100800ffff00000000000045100200000034 ,
                        0x100000031008000300010004000400331000005110080000010200000002000d ,
                        0x1009000000064578747261735110080001010200000002005110080002010200 ,
                        0x0000000006100800ffff01000100000045100200000034100000031008000300 ,
                        0x010008000800331000005110080000010200000003000d100d0000000a477261 ,
                        0x6e64546f74616c51100800010102000000030051100800020102000000000006 ,
                        0x100800ffff020002000000451002000000341000004410030009000046100200 ,
                        0x0100411012000000cc000000a2020000970d0000570b0000331000004f101400 ,
                        0x02000200060000002c020000600e0000030d00001d1012000000000000000000 ,
                        0x00000000000000000000331000001e101a000200030100000000000000000000 ,
                        0x000000000000000000002300341000001d101200010000000000000000000000 ,
                        0x000000000000331000001f102a00000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000001e011e101a0002000301 ,
                        0x0000000000000000000000000000000000000000230034100000141014000000 ,
                        0x00000000000000000000000000000000000033100000171006009cff96000200 ,
                        0x22100a0032100400000003003310000007100a00000000000000000009000a10 ,
                        0x0c00ffffff000000000000000000341000003410000024100200020025101a00 ,
                        0x0202010000000000ddffffffc7ffffff0000000000000000b100331000004f10 ,
                        0x1400020002000000000000000000000000000000000026100200020051100800 ,
                        0x00010200ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xbd3b00001b250000540f0000010009000003aa07000007001c00000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c0267014202030000001e00040000002701ffff050000000b02 ,
                        0x0000000001000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000001000 ,
                        0x0000fb02f5ff000000000000bc020000000000000022417269616c00a8810400 ,
                        0x00002d01000010000000fb021000070000000000bc0200000000010202225379 ,
                        0x7374656d006e040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016046701 ,
                        0x420200000000040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000ff00070000001604400112023f002100040000002d01040004000000 ,
                        0x2d010600040000002d010200040000002701ffff030000001e00040000002d01 ,
                        0x0100040000002d010500040000002d0103000500000001028020600205000000 ,
                        0x0902ffffff00070000001604680142020000000009000000fa02000001000000 ,
                        0x000000002200040000002d01070004000000f001030005000000140240002200 ,
                        0x050000000102ffffff000400000004010d000400000002010100050000001302 ,
                        0x3e0122000500000014023e011f000500000013023e0122000500000014022201 ,
                        0x1f000500000013022201220005000000140206011f0005000000130206012200 ,
                        0x050000001402e9001f00050000001302e9002200050000001402cd001f000500 ,
                        0x00001302cd002200050000001402b1001f00050000001302b100220005000000 ,
                        0x140295001f000500000013029500220005000000140278001f00050000001302 ,
                        0x780022000500000014025c001f000500000013025c0022000500000014024000 ,
                        0x1f00050000001302400022000500000014023e0122000500000013023e011002 ,
                        0x050000001402410122000500000013023e01220005000000140241019e000500 ,
                        0x000013023e019e00050000001402410119010500000013023e01190105000000 ,
                        0x1402410195010500000013023e01950105000000140241011002050000001302 ,
                        0x3e011002040000002d010400040000002d010600040000002d01020004000000 ,
                        0x2701ffff030000001e00040000002d010100040000002d010500040000002d01 ,
                        0x0700050000000102ffffff00050000000902ffffff0007000000160426004f01 ,
                        0x0b00f300050000000902000000000400000004010d0004000000020101001c00 ,
                        0x0000320a1100f9000e0000004f766572616c6c2053636f726573080006000700 ,
                        0x05000600030003000300070006000700050007000700040000002d0104000400 ,
                        0x00002d010600040000002d010200040000002701ffff030000001e0004000000 ,
                        0x2d010100040000002d010500040000002d010700050000000102ffffff000500 ,
                        0x00000902000000000700000016046801420200000000040000002d0104000400 ,
                        0x00002d010600040000002d010200040000002701ffff030000001e0004000000 ,
                        0x2d010100040000002d010500040000002d010700050000000102ffffff000500 ,
                        0x00000902000000000700000016046801420200000000040000002d0104000400 ,
                        0x00002d010600040000002d010200040000002701ffff030000001e0004000000 ,
                        0x2d010100040000002d010500040000002d010700050000000102ffffff000500 ,
                        0x000009020000000007000000160469014302ffffffff0400000004010d000400 ,
                        0x00000201010009000000320a3701140001000000300006000a000000320a1b01 ,
                        0x0e00020000003530060006000c000000320aff00080003000000313030000600 ,
                        0x060006000c000000320ae200080003000000313530000600060006000c000000 ,
                        0x320ac600080003000000323030000600060006000c000000320aaa0008000300 ,
                        0x0000323530000600060006000c000000320a8e00080003000000333030000600 ,
                        0x060006000c000000320a7100080003000000333530000600060006000c000000 ,
                        0x320a5500080003000000343030000600060006000c000000320a390008000300 ,
                        0x000034353000060006000600040000002d010400040000002d01060004000000 ,
                        0x2d010200040000002701ffff030000001e00040000002d010100040000002d01 ,
                        0x0500040000002d010700050000000102ffffff00050000000902000000000700 ,
                        0x000016046801420200000000040000002d010400040000002d01060004000000 ,
                        0x2d010200040000002701ffff030000001e00040000002d010100040000002d01 ,
                        0x0500040000002d010700050000000102ffffff00050000000902000000000700 ,
                        0x0000160469014302ffffffff0400000004010d0004000000020101000f000000 ,
                        0x320a47014f00050000004a5544414800060007000700080007000f000000320a ,
                        0x4701ca00050000004153484552000800070007000600070012000000320a4701 ,
                        0x3f01070000004550485241494d000600070007000700080003000a000d000000 ,
                        0x320a4701c601040000004c4556490700060008000300040000002d0104000400 ,
                        0x00002d010600040000002d010200040000002701ffff030000001e0004000000 ,
                        0x2d010100040000002d010500040000002d010700050000000102ffffff000500 ,
                        0x0000090200000000070000001604680142020000000007000000fc0201000000 ,
                        0x00000000040000002d01030004000000f00105000400000004010d0004000000 ,
                        0x02010200070000001b043b003d0214000802040000002d010400040000002d01 ,
                        0x0600040000002d010200040000002701ffff030000001e00040000002d010100 ,
                        0x040000002d010300040000002d010700050000000102ffffff00050000000902 ,
                        0x000000000700000016043a003d0214000802040000002d010400040000002d01 ,
                        0x0600040000002d010200040000002701ffff030000001e00040000002d010100 ,
                        0x040000002d010300040000002d010700050000000102ffffff00050000000902 ,
                        0x000000000700000016043a003d021400080209000000fa020000010000000000 ,
                        0x00002200040000002d01050004000000f001070007000000fc02000080206002 ,
                        0xffff040000002d01070004000000f0010300050000000902ffffff0005000000 ,
                        0x0102802060020400000004010d000400000002010200070000001b0422001302 ,
                        0x1b000c0205000000090200000000050000000102ffffff000400000002010100 ,
                        0x04000000030000001e00040000002d010000040000002d010300040000002d01 ,
                        0x0600050000000102ffffff000500000009020000000007000000160467014202 ,
                        0x00000000040000002d010400040000002d010500040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000040000002d010300040000002d01 ,
                        0x0600050000000102ffffff000500000009020000000007000000160468014302 ,
                        0xffffffff0400000004010d0004000000020101000d000000320a460131000400 ,
                        0x00004341524d0800080007000a000a000000320a460178000200000043480800 ,
                        0x07000d000000320a4601b000040000004348495308000700030007000c000000 ,
                        0x320a4601ee0003000000434f43000800080008000c000000320a46012e010300 ,
                        0x00004a5043000600070008000c000000320a46016b0103000000535043000700 ,
                        0x070008000c000000320a4601a7010300000053544d00070007000a000d000000 ,
                        0x320a4601e501040000005452494e0700070003000700040000002d0104000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010300040000002d010600050000000102ffffff000500 ,
                        0x0000090200000000070000001604670142020000000007000000fc0201000000 ,
                        0x00000000040000002d01020004000000f00103000400000004010d0004000000 ,
                        0x02010200070000001b043b003d0214000802040000002d010400040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010200040000002d010600050000000102ffffff00050000000902 ,
                        0x000000000700000016043a003d0214000802040000002d010400040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010200040000002d010600050000000102ffffff00050000000902 ,
                        0x000000000700000016043a003d021400080209000000fa020000010000000000 ,
                        0x00002200040000002d01030004000000f001060007000000fc02000080206000 ,
                        0xffff040000002d01060004000000f0010200050000000902ffffff0005000000 ,
                        0x0102802060000400000004010d000400000002010200070000001b0422001302 ,
                        0x1b000c0205000000090200000000050000000102ffffff000400000002010100 ,
                        0x040000002e01180010000000320a220016020600000045787472617306000600 ,
                        0x0400050006000700040000002e010000040000002d010400040000002d010500 ,
                        0x040000002d010100040000002701ffff030000001e00040000002d0100000400 ,
                        0x00002d010600040000002d010300050000000102ffffff000500000009020000 ,
                        0x00000700000016043a003d0214000802040000002d010400040000002d010500 ,
                        0x040000002d010100040000002701ffff030000001e00040000002d0100000400 ,
                        0x00002d010600040000002d010300050000000102ffffff000500000009020000 ,
                        0x00000700000016043a003d021400080207000000fc0200008080ff00ffff0400 ,
                        0x00002d01020004000000f0010600050000000902ffffff000500000001028080 ,
                        0xff000400000004010d000400000002010200070000001b04350013022e000c02 ,
                        0x05000000090200000000050000000102ffffff00040000000201010004000000 ,
                        0x2e01180010000000320a3500160206000000506f696e74730700070003000700 ,
                        0x04000700040000002e010000040000002d010400040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d010300050000000102ffffff00050000000902000000000700 ,
                        0x000016043a003d0214000802040000002d010400040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d010300050000000102ffffff00050000000902000000000700 ,
                        0x00001604670142020000000007000000fc020000000000000000040000002d01 ,
                        0x060004000000f0010200040000002d01040004000000f0010300040000002701 ,
                        0xffff050000000c0267014202030000001e00050000000102ffffff0005000000 ,
                        0x090200002e01180010000000320a220016020600000045787472617306000600 ,
                        0x0400050006000700040000002e010000040000002d010400040000002d010600 ,
                        0x040000002d010200040000002701ffff030000001e00040000002d0101000400 ,
                        0x00002d010700040000002d010500050000000102ffffff000500000009020000 ,
                        0x00000700000016043a003d0214000802040000002d010400040000002d010600 ,
                        0x040000002d010200040000002701ffff030000001e00040000002d0101000400 ,
                        0x00002d010700040000002d010500050000000102ffffff000500000009020000 ,
                        0x00000700000016043a003d021400080207000000fc0200008080ff02ffff0400 ,
                        0x00002d01030004000000f0010700050000000902ffffff000500000001028080 ,
                        0xff020400000004010d000400000002010200070000001b04350013022e000c02 ,
                        0x05000000090200000000050000000102ffffff00040000000201010004000000 ,
                        0x2e01180010000000320a3500160206000000506f696e74730700070003000700 ,
                        0x04000700040000002e010000040000002d010400040000002d01060004000000 ,
                        0x2d010200040000002701ffff030000001e00040000002d010100040000002d01 ,
                        0x0300040000002d010500050000000102ffffff00050000000902000000000700 ,
                        0x000016043a003d0214000802040000002d010400040000002d01060004000000 ,
                        0x2d010200040000002701ffff030000001e00040000002d010100040000002d01 ,
                        0x0300040000002d010500050000000102ffffff00050000000902000000000700 ,
                        0x00001604680142020000000007000000fc020000000000000000040000002d01 ,
                        0x070004000000f0010300040000002d01040004000000f0010500040000002701 ,
                        0xffff050000000c0268014202030000001e00050000000102ffffff0005000000 ,
                        0x090200000000040000002701ffff050000000b0200000000030000001e000500 ,
                        0x00000102ffffff0005000000090200000000040000002701ffff030000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW [House points - GrandTotal].H_Code, Sum([House points - Grand"
                        "Total].SumOfPoints) AS Points, Sum([House points - GrandTotal].Extras) AS Extras"
                        " FROM [House points - GrandTotal] GROUP BY [House points - GrandTotal].H_Code;"
                    Class ="MSGraph.Chart.5"
                    OLEClass ="Microsoft Graph 5.0"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter1"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

Dim sHTML As String, rHTML As String, PageNum As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String
Dim GenerateHTML As Boolean
Dim ExportOleChart As Boolean

Const ReportTitle = "Overall Statistical Summary - Ordered by Grand Total"
Const repName = "over" ' Keep to 4 letters or less (and unique from all other reports

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    Dim BGcolor As String

    If GenerateHTML And Not IsNull(Me!Place) Then
    
        DetailCount = DetailCount + 1
        
        If DetailCount = 1 Then

            If PageNum Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
            
            ' *** Create general record header ***
            Call RowStart(rHTML)
            
            Call CellStart(rHTML, "Left", "Center", "10%", cCream, 1)
            Call SpaceIndent(rHTML, 2)
            Call Text(rHTML, "<B>", "</B>", "PLACE")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "40%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "TEAM")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "TOTAL")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "EXTRAS")
            Call CellEnd(rHTML)
            
            
            Call CellStart(rHTML, "Center", "Center", "15%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "GRAND TOT.")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "15%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "% TOTAL")
            Call CellEnd(rHTML)
            
        End If
        
        If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
        'If (Me!Place = 1) And (Me!f_lev = 0) Then BGcolor = cLightRed

        Call RowStart(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call SpaceIndent(rHTML, 2)
        Call Text(rHTML, "", "", Me!Place)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_NAme)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!SumOfPoints)
        Call CellEnd(rHTML)
        
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Extras)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "15%", BGcolor, 1)
        Call Text(rHTML, "<B>", "</B>", Me!GrandTotal)
        Call CellEnd(rHTML)
        
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!PercentileTotal)
        Call CellEnd(rHTML)
        
        
        Call RowEnd(rHTML)
    End If


End Sub

Private Sub PageFooter2_Format(Cancel As Integer, FormatCount As Integer)

'On Error Resume Next

  If GenerateHTML Then
    
    Template = DLookup("[TemplateFile]", "MiscHTML")
    Call TableEnd(rHTML)
    
    
    Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & PageNum, ReportHead, repName)
  
  End If

End Sub

Private Sub PageHeader0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    If GenerateHTML Then
        DetailCount = 0
        PageNum = PageNum + 1
        rHTML = ""
        
        If PageNum > 1 Then
            PrevPage = Link(repName & PageNum - 1 & ".htm", "Previous Page")
        Else
            PrevPage = ""
        End If
        NextPage = Link(repName & PageNum + 1 & ".htm", "Next Page")
        
        Call DivOpen(rHTML, "header")
        Call TableOpen(rHTML, "header-" & repName)
            
        Call Cell(rHTML, Heading(3, "Overall Statistical Summary - Ordered by Grand Total", 3), "titleTable")
        Call TableEnd(rHTML)
        Call DivClose(rHTML)

        Call TableOpen(rHTML, "data-" & repName)

    End If



End Sub


Private Sub Report_Close()
On Error GoTo Report_Close_Err

  If ExportOleChart Then
    Dim oleGraph As Object
    HTMLFileLocation = DLookup("[HTMLlocation]", "MiscHTML")
    fileLocation = HTMLFileLocation & "\overall.jpg"
    Set oleGraph = Me.oleChart.Object
    
    oleGraph.Export FileName:=fileLocation
    
    oleGraph.Close
    Set oleGraph = Nothing
    
  End If
  
Report_Close_Exit:
  DoCmd.RunMacro "ReportPopup-Update"
  Exit Sub
  
Report_Close_Err:
  GoTo Report_Close_Exit
  
End Sub

Private Sub Report_Open(Cancel As Integer)

On Error Resume Next

    GenerateHTML = GlobalGenerateHTML
    ExportOleChart = GlobalGenerateHTML
    If GenerateHTML Then
        PleaseWaitMsg = "Generating HTML for """ & ReportTitle & """.  Please wait..."
        DoCmd.RunMacro "ShowPleaseWait"
    End If

    
    PageNum = 0
    LastPage = False
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")

End Sub

Private Sub ReportFooter1_Format(Cancel As Integer, FormatCount As Integer)
On Error Resume Next

    If GenerateHTML Then
        GenerateHTML = False
        
        NextPage = ""
        Call TableEnd(rHTML)
        rHTML = rHTML & " <p align=""center""><img border=""0"" src=""overall.jpg"" </p>"
        Template = DLookup("[TemplateFile]", "MiscHTML")
        'TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & PageNum, ReportHead, repName)
        
        DoCmd.RunMacro "ClosePleaseWait"
        
    End If


End Sub
