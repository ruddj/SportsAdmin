Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =119
    Left =3810
    Top =360
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x2d423b736dd7e240
    End
    RecordSource ="HousePoints-Total-Sex-F"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    PrtMip = Begin
        0x370200003702000045020000d002000000000000a02900005401000001000000 ,
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="HousePoints-Total-Sex.Sex Sub"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SumOfPoints"
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
            Height =1247
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
                    Width =5880
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="Overall Statistical Summary - by Gender"
                    FontName ="times New Roman"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1045
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =1247
                    Top =685
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
                    Left =5669
                    Top =675
                    Width =1365
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
                    Left =8502
                    Top =623
                    Width =1455
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Grand Total"
                    FontName ="times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Left =566
                    Top =1020
                    Width =10077
                    Name ="Line112"
                End
                Begin Label
                    TextFontFamily =18
                    Left =56
                    Top =113
                    Width =900
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text114"
                    Caption ="SEX:"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =963
                    Top =170
                    Width =3576
                    Height =330
                    FontSize =11
                    FontWeight =700
                    Name ="Age"
                    ControlSource ="HousePoints-Total-Sex.Sex Sub"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =355
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =1247
                    Width =4296
                    Height =330
                    FontSize =10
                    FontWeight =700
                    Name ="H_NAme"
                    ControlSource ="H_NAme"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5672
                    Width =1011
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Field100"
                    ControlSource ="H_Code"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =7941
                    Width =1521
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Field106"
                    ControlSource ="SumOfPoints"

                End
                Begin Line
                    Left =510
                    Top =340
                    Width =10077
                    Name ="Line86"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =510
                    Width =636
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Place"
                    ControlSource ="=1"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =170
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =5832
            OnFormat ="[Event Procedure]"
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =5442
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =2
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =5442
                    Width =9471
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
                    Top =5442
                    Width =10596
                    Name ="Line87"
                End
                Begin Chart
                    ColumnHeads = NotDefault
                    Locked = NotDefault
                    SizeMode =3
                    RowSourceTypeInt =2
                    Left =226
                    Top =115
                    Width =10205
                    Height =5043
                    TabIndex =1
                    Name ="oleChart"
                    OleData = Begin
                        0x003c0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
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
                        0x00000046000000000000000000000000000000000000000007000000400a0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000063000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002800000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000300000094080000 ,
                        0x0000000005000000fdfffffffffffffffffffffffffffffffefffffffeffffff ,
                        0x10000000090000000a0000000b0000000c00000019000000ffffffffffffffff ,
                        0x180000001100000017000000ffffffffffffffffffffffffffffffffffffffff ,
                        0x0f000000feffffff1a0000001b0000001c000000feffffffffffffffffffffff ,
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
                        0x0000000016000500ffffffffffffffff050000000308020000000000c0000000 ,
                        0x00000046000000000000000000000000a05994c1ce20bf010e000000400a0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000068000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002800000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000300000094080000 ,
                        0x00000000ffffffffffffffff05000000fdfffffffefffffffeffffffffffffff ,
                        0xffffffff090000000a0000000b0000000c00000019000000feffffff10000000 ,
                        0x0d0000001100000017000000ffffffffffffffffffffffffffffffffffffffff ,
                        0x0f000000ffffffff1a0000001b0000001c000000feffffffffffffffffffffff ,
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
                        0xfffffffffeffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x200000002100000022000000230000002400000025000000feffffffffffffff ,
                        0xfffffffffeffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
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
                        0x0000000000000000000000000000000000000000000000000800000008100000 ,
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
                        0x200000002100000022000000230000002400000025000000feffffffffffffff ,
                        0xfffffffffeffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
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
                        0x000000000908080080050500e209c9071e041600050013222422232c2323303b ,
                        0x5c2d222422232c2323301e041b00060018222422232c2323303b5b5265645d5c ,
                        0x2d222422232c2323301e041c00070019222422232c2323302e30303b5c2d2224 ,
                        0x22232c2323302e30301e04210008001e222422232c2323302e30303b5b526564 ,
                        0x5d5c2d222422232c2323302e30301e0433002a00305f2d2224222a20232c2323 ,
                        0x305f2d3b5c2d2224222a20232c2323305f2d3b5f2d2224222a20222d225f2d3b ,
                        0x5f2d405f2d1e042a002900275f2d2a20232c2323305f2d3b5c2d2a20232c2323 ,
                        0x305f2d3b5f2d2a20222d225f2d3b5f2d405f2d1e043b002c00385f2d2224222a ,
                        0x20232c2323302e30305f2d3b5c2d2224222a20232c2323302e30305f2d3b5f2d ,
                        0x2224222a20222d223f3f5f2d3b5f2d405f2d1e0432002b002f5f2d2a20232c23 ,
                        0x23302e30ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0x48460000ba220000e00f0000010009000003f007000007002400000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c025001a802030000001e00040000002701ffff050000000b02 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000001000 ,
                        0x0000fb02f5ff000000000000bc020000000000000022417269616c00ab810400 ,
                        0x00002d01000010000000fb021000070000000000bc0200000000010202225379 ,
                        0x7374656d006e040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016045001 ,
                        0xa80200000000040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016045001 ,
                        0xa8020000000007000000150439009b0213006e02040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000050000000102ffffff0005000000 ,
                        0x0902000000000700000016045001a80200000000040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000050000000102ffffff0005000000 ,
                        0x0902000000000700000016044c01a40204000400040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000050000000102ffffff0005000000 ,
                        0x09020000000007000000160428017e023e002100040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000050000000102ffffff0005000000 ,
                        0x09020000000007000000160427017f023c00210009000000fa02000001000000 ,
                        0x000000002200040000002d01020007000000fc0200008080ff00ffff04000000 ,
                        0x2d010300050000000902ffffff000500000001028080ff000400000004010d00 ,
                        0x04000000020102000e000000240305003200c6004800c6004800260132002601 ,
                        0x3200c6000e000000240305007d00d9009300d900930026017d0026017d00d900 ,
                        0x0e00000024030500c9009600df009600df002601c9002601c90096000e000000 ,
                        0x240305001401ec002a01ec002a012601140126011401ec000e00000024030500 ,
                        0x5f01000175010001750126015f0126015f0100010e00000024030500aa01b300 ,
                        0xc001b300c0012601aa012601aa01b3000e00000024030500f601e3000c02e300 ,
                        0x0c022601f6012601f601e3000e000000240305004102f6005702f60057022601 ,
                        0x410226014102f60007000000fc02000080206000ffff040000002d0104000400 ,
                        0x0000f0010300050000000102802060000e0000002403050048009f005d009f00 ,
                        0x5d0026014800260148009f000e000000240305009300ba00a900ba00a9002601 ,
                        0x930026019300ba000e00000024030500df005c00f4005c00f4002601df002601 ,
                        0xdf005c000e000000240305002a01d5003f01d5003f0126012a0126012a01d500 ,
                        0x0e000000240305007501f0008a01f0008a012601750126017501f0000e000000 ,
                        0x24030500c0018400d6018400d6012601c0012601c00184000e00000024030500 ,
                        0x0c02c8002102c800210226010c0226010c02c8000e000000240305005702e300 ,
                        0x6c02e3006c022601570226015702e30009000000fa0200000000000000000000 ,
                        0x2200040000002d01030007000000fc020000ffffff000000040000002d010500 ,
                        0x040000002d010100040000002701ffff030000001e00040000002d0100000400 ,
                        0x00002d010400040000002d01020005000000010280206000050000000902ffff ,
                        0xff0007000000160428017e023e002100040000002d010300040000002d010500 ,
                        0x040000002d010100040000002701ffff030000001e00040000002d0100000400 ,
                        0x00002d010400040000002d01020005000000010280206000050000000902ffff ,
                        0xff000700000016045001a8020000000009000000fa0200000100000000000000 ,
                        0x2200040000002d01060004000000f00102000500000014023f00220005000000 ,
                        0x0102ffffff000400000004010d00040000000201010005000000130226012200 ,
                        0x05000000140226011f000500000013022601220005000000140200011f000500 ,
                        0x0000130200012200050000001402d9001f00050000001302d900220005000000 ,
                        0x1402b3001f00050000001302b30022000500000014028c001f00050000001302 ,
                        0x8c00220005000000140266001f00050000001302660022000500000014023f00 ,
                        0x1f000500000013023f0022000500000014022601220005000000130226017c02 ,
                        0x050000001402290122000500000013022601220005000000140229016d000500 ,
                        0x0000130226016d000500000014022901b9000500000013022601b90005000000 ,
                        0x1402290104010500000013022601040105000000140229014f01050000001302 ,
                        0x26014f0105000000140229019a0105000000130226019a010500000014022901 ,
                        0xe6010500000013022601e6010500000014022901310205000000130226013102 ,
                        0x05000000140229017c0205000000130226017c02040000002d01030004000000 ,
                        0x2d010500040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000040000002d010400040000002d010600050000000102ffffff0005000000 ,
                        0x0902ffffff0007000000160426008e010b001a01050000000902000000000400 ,
                        0x000004010d00040000000201010024000000320a1100200113000000546f7461 ,
                        0x6c20506f696e7473206279205365780007000700040006000300030007000700 ,
                        0x03000700040007000300070006000300070007000600040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016045001a80200000000040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016045001a80200000000040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016045101a902ffffffff0400000004010d000400 ,
                        0x00000201010009000000320a1f01140001000000300006000a000000320af900 ,
                        0x0e00020000003230060006000a000000320ad2000e0002000000343006000600 ,
                        0x0a000000320aac000e00020000003630060006000a000000320a85000e000200 ,
                        0x00003830060006000c000000320a5f0008000300000031303000060006000600 ,
                        0x0c000000320a380008000300000031323000060006000600040000002d010300 ,
                        0x040000002d010500040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010400040000002d010600050000000102ffffff00 ,
                        0x0500000001000002040000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000020000000a0000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000038000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0308020000000000c000000000000046 ,
                        0x190000004d6963726f736f667420477261706820393720436861727400070000 ,
                        0x0047426966663500100000004d5347726170682e43686172742e3800f439b271 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000908080080050500e209c9071e041600050013222422232c2323303b ,
                        0x5c2d222422232c2323301e041b00060018222422232c2323303b5b5265645d5c ,
                        0x2d222422232c2323301e041c00070019222422232c2323302e30303b5c2d2224 ,
                        0x22232c2323302e30301e04210008001e222422232c2323302e30303b5b526564 ,
                        0x5d5c2d222422232c2323302e30301e0433002a00305f2d2224222a20232c2323 ,
                        0x305f2d3b5c2d2224222a20232c2323305f2d3b5f2d2224222a20222d225f2d3b ,
                        0x5f2d405f2d1e042a002900275f2d2a20232c2323305f2d3b5c2d2a20232c2323 ,
                        0x305f2d3b5f2d2a20222d225f2d3b5f2d405f2d1e043b002c00385f2d2224222a ,
                        0x20232c2323302e30305f2d3b5c2d2224222a20232c2323302e30305f2d3b5f2d ,
                        0x2224222a20222d223f3f5f2d3b5f2d405f2d1e0432002b002f5f2d2a20232c23 ,
                        0x23302e300000e2ffffffc3ffffff0000000000000000b100331000004f101400 ,
                        0x0200020000000000000000000000000000000000511008000001020000000000 ,
                        0x3410000032100400000003003310000007100a00000000000000000009000a10 ,
                        0x0c00ffffff000000000000000000341000003410000024100200020025101a00 ,
                        0x0202010000000000e2ffffffc3ffffff0000000000000000b100331000004f10 ,
                        0x1400020002000000000000000000000000000000000026100200020051100800 ,
                        0x000102000000000034100000341000003410000025101a000202010000000000 ,
                        0x7c06000056000000a9020000330100008100331000004f101400020002000000 ,
                        0x00000000000072000000190000005110080000010200000000000d1016000000 ,
                        0x13546f74616c20506f696e747320627920536578271006000100000000003410 ,
                        0x00003410000000000a00000006000000020000000a0000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003410000025101a0002020100000000007c06000056000000a9020000 ,
                        0x330100008100331000004f101400020002000000000000000000720000001900 ,
                        0x00005110080000010200000000000d101600000013546f74616c20506f696e74 ,
                        0x732062792053657827100600010000000000341000003410000000000a000000 ,
                        0x10000000305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a20222d223f3f ,
                        0x5f2d3b5f2d405f2d1e041a00a40017222422232c2323305f293b5c2822242223 ,
                        0x2c2323305c291e041f00a5001c222422232c2323305f293b5b5265645d5c2822 ,
                        0x2422232c2323305c291e042000a6001d222422232c2323302e30305f293b5c28 ,
                        0x222422232c2323302e30305c291e042500a70022222422232c2323302e30305f ,
                        0x293b5b5265645d5c28222422232c2323302e30305c291e043500a800325f2822 ,
                        0x24222a20232c2323305f293b5f282224222a205c28232c2323305c293b5f2822 ,
                        0x24222a20222d225f293b5f28405f291e042c00a900295f282a20232c2323305f ,
                        0x293b5f282a205c28232c2323305c293b5f282a20222d225f293b5f28405f291e ,
                        0x043d00aa003a5f282224222a20232c2323302e30305f293b5f282224222a205c ,
                        0x28232c2323302e30305c293b5f282224222a20222d223f3f5f293b5f28405f29 ,
                        0x1e043400ab00315f282a20232c2323302e30305f293b5f282a205c28232c2323 ,
                        0x302e30305c293b5f282a20222d223f3f5f293b5f28405f2931001400a0000100 ,
                        0xff7fbc0200000000007205417269616c31001400c8000100ff7fbc0200000000 ,
                        0x007205417269616c31001400a0000100ff7fbc0200000002007205417269616c ,
                        0x31001400c8000100ff7fbc0200000000007205417269616c3d00120092040c03 ,
                        0x50284515000000003e200000000085000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x003d002610020003005310040000000400541004000000030055100600000000 ,
                        0x00000104000e000000000000000006485f436f646504000c0000000100000000 ,
                        0x04426f797304000d0000000200000000054769726c7304000b00010000000000 ,
                        0x0003434f4303000f0001000100000000000000000000304003000f0001000200 ,
                        0x000000000000000000204004000b0002000000000000034a504303000f000200 ,
                        0x01000000000000000000002c4003000f00020002000000000000000000001c40 ,
                        0x04000c0003000000000000045452494e03000f00030001000000000000000000 ,
                        0x00284003000f00030002000000000000000000001840571001000159100800b4 ,
                        0x002c01942f8a1b3d000a0000000000f627000f00003e000e0001010100010100 ,
                        0x010001000000005810020000001d001100030100010000000100010001000100 ,
                        0x010034100000011002000000021010000000000000000000e8fffd010000fc00 ,
                        0x33100000a0000400010001000310080003000100030003003310000051100800 ,
                        0x00010200000001000d100700000004426f797351100800010102000000010051 ,
                        0x100800020102000000000006100800ffff000000000000451002000000341000 ,
                        0x000310080025101a000202010000000000e2ffffffc3ffffff00000000000000 ,
                        0x00b100331000004f101400020002000000000000000000000000000000000051 ,
                        0x10080000010200000000003410000032100400000003003310000007100a0000 ,
                        0x0000000000000009000a100c00ffffff00000000000000000034100000341000 ,
                        0x0024100200020025101a000202010000000000e2ffffffc3ffffff0000000000 ,
                        0x000000b100331000004f10140002000200000000000000000000000000000000 ,
                        0x0026100200020051100800000102000000000034100000341000003410000025 ,
                        0x101a0002020100000000007c06000056000000a9020000330100008100331000 ,
                        0x004f101400020002000000000000000000720000001900000051100800000102 ,
                        0x00000000000d101600000013546f74616c20506f696e74732062792053657827 ,
                        0x100600010000000000341000003410000000000a00000006000000030000000a ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003410000025101a0002020100000000007c06000056000000a9020000 ,
                        0x330100008100331000004f101400020002000000000000000000720000001900 ,
                        0x00005110080000010200000000000d101600000013546f74616c20506f696e74 ,
                        0x732062792053657827100600010000000000341000003410000000000a000000 ,
                        0x10000000305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a20222d223f3f ,
                        0x5f2d3b5f2d405f2d1e041a00a40017222422232c2323305f293b5c2822242223 ,
                        0x2c2323305c291e041f00a5001c222422232c2323305f293b5b5265645d5c2822 ,
                        0x2422232c2323305c291e042000a6001d222422232c2323302e30305f293b5c28 ,
                        0x222422232c2323302e30305c291e042500a70022222422232c2323302e30305f ,
                        0x293b5b5265645d5c28222422232c2323302e30305c291e043500a800325f2822 ,
                        0x24222a20232c2323305f293b5f282224222a205c28232c2323305c293b5f2822 ,
                        0x24222a20222d225f293b5f28405f291e042c00a900295f282a20232c2323305f ,
                        0x293b5f282a205c28232c2323305c293b5f282a20222d225f293b5f28405f291e ,
                        0x043d00aa003a5f282224222a20232c2323302e30305f293b5f282224222a205c ,
                        0x28232c2323302e30305c293b5f282224222a20222d223f3f5f293b5f28405f29 ,
                        0x1e043400ab00315f282a20232c2323302e30305f293b5f282a205c28232c2323 ,
                        0x302e30305c293b5f282a20222d223f3f5f293b5f28405f2931001400a0000100 ,
                        0xff7fbc0200000000006105417269616c31001400c8000100ff7fbc0200000000 ,
                        0x006105417269616c31001400a0000100ff7fbc0200000002006105417269616c ,
                        0x31001400c8000100ff7fbc0200000000006105417269616c3d00120092040c03 ,
                        0x50284515000000003e200000000085000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x003d002610020003005310040000000400541004000000040055100600000000 ,
                        0x00000104000e000000000000000006485f436f646504000a0000000100000000 ,
                        0x023c3e04000c000000020000000004426f797304000d00000003000000000547 ,
                        0x69726c7304000b000100000000000003434f4303000f00010002000000000000 ,
                        0x00000000304003000f0001000300000000000000000000204004000b00020000 ,
                        0x00000000034a504303000f00020002000000000000000000002c4003000f0002 ,
                        0x0003000000000000000000001c4004000c0003000000000000045452494e0300 ,
                        0x0f0003000200000000000000000000284003000f000300030000000000000000 ,
                        0x00001840571001000159100800b4002c01942f8a1b3d000a0000000000f62700 ,
                        0x0f00003e000e0001010100010100010001000000005810020000001d00110003 ,
                        0x0100010000000100010001000100010034100000011002000000021010000000 ,
                        0x000000000000e8fffd010000fc0033100000a000040001000100031008000300 ,
                        0x010003000300331000005110080000010200000001000d1005000000023c3e51 ,
                        0x100800010102000000010051100800020102000000000006100800ffff000000 ,
                        0x0000004510020000003410000003100800030001000300030033100000511008 ,
                        0x0000010200000002000d100700000004426f7973511008000101020000000200 ,
                        0x51100800020102000000000006100800ffff0100010000004510020000003410 ,
                        0x0000031008000300010003000300331000005110080000010200000003000d10 ,
                        0x08000000054769726c7351100800010102000000030051100800020102000000 ,
                        0x000006100800ffff020002000000451002000000341000004410030009000046 ,
                        0x100200010041101200000089000000c80200002e0e0000120b0000331000004f ,
                        0x101400020002000500000048020000b10e0000de0c00001d1012000000000000 ,
                        0x00000000000000000000000000331000001e101a000200030100000000000000 ,
                        0x000000000000000000000000002300341000001d101200010000000000000000 ,
                        0x000000000000000000331000001f102a00000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000001f011e101a0002 ,
                        0x0003010000000000000000000000000000000000000000230034100000141014 ,
                        0x0000000000000000000000000000000000000000003310000017100600000096 ,
                        0x00000022100a0000000000000000000f0015101400640e0000ac0000000d0100 ,
                        0x00bb02000007011200331000004f10140005000200640e0000ad000000000000 ,
                        0x0000000001000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000020000000a0000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000300010003000300331000005110080000010200000002000d1008 ,
                        0x000000054769726c735110080001010200000002005110080002010200000000 ,
                        0x0006100800ffff01000100000045100200000034100000031008000300010003 ,
                        0x000300331000005110080000010200000003000d1008000000054769726c7351 ,
                        0x100800010102000000030051100800020102000000000006100800ffff020002 ,
                        0x0000004510020000003410000044100300090000461002000100411012000000 ,
                        0x89000000c80200002e0e0000120b0000331000004f1014000200020005000000 ,
                        0x48020000b10e0000de0c00001d10120000000000000000000000000000000000 ,
                        0x0000331000001e101a0002000301000000000000000000000000000000000000 ,
                        0x00002300341000001d1012000100000000000000000000000000000000003310 ,
                        0x00001f102a000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000001f011e101a00020003010000000000000000 ,
                        0x0000000000000000000000002300341000001410140000000000000000000000 ,
                        0x00000000000000000000331000001710060000009600000022100a0000000000 ,
                        0x000000000f0015101400640e0000ac0000000d010000d2010000070112003310 ,
                        0x00004f10140005000200640e0000ad000000000000000000000025101a000202 ,
                        0x0100000001000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000020000000a0000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000902000000000700000016045001a80200000000040000002d010300 ,
                        0x040000002d010500040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010400040000002d010600050000000102ffffff00 ,
                        0x050000000902000000000700000016045101a902ffffffff0400000004010d00 ,
                        0x04000000020101000d000000320a2f013800040000004341524d080008000700 ,
                        0x0a000a000000320a2f018c00020000004348080007000d000000320a2f01d200 ,
                        0x040000004348495308000700030007000c000000320a2f011d0103000000434f ,
                        0x43000800080008000c000000320a2f016b01030000004a504300060007000800 ,
                        0x0c000000320a2f01b50103000000535043000700070008000c000000320a2f01 ,
                        0xff010300000053544d00070007000a000d000000320a2f014a02040000005452 ,
                        0x494e0700070003000700040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010400 ,
                        0x040000002d010600050000000102ffffff000500000009020000000007000000 ,
                        0x16045001a8020000000007000000fc020100000000000000040000002d010200 ,
                        0x04000000f00104000400000004010d000400000002010200070000001b043a00 ,
                        0x9b0213006e02040000002d010300040000002d010500040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01020004000000 ,
                        0x2d010600050000000102ffffff00050000000902000000000700000016043900 ,
                        0x9b0213006e02040000002d010300040000002d010500040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01020004000000 ,
                        0x2d010600050000000102ffffff00050000000902000000000700000016043900 ,
                        0x9b0213006e0209000000fa02000001000000000000002200040000002d010400 ,
                        0x04000000f001060007000000fc0200008080ff00ffff040000002d0106000400 ,
                        0x0000f0010200050000000902ffffff000500000001028080ff00040000000401 ,
                        0x0d000400000002010200070000001b04210079021a0072020500000009020000 ,
                        0x0000050000000102ffffff000400000002010100040000002e0118000d000000 ,
                        0x320a21007c0204000000426f79730700070006000700040000002e0100000400 ,
                        0x00002d010300040000002d010500040000002d010100040000002701ffff0300 ,
                        0x00001e00040000002d010000040000002d010600040000002d01040005000000 ,
                        0x0102ffffff000500000009020000000007000000160439009b0213006e020400 ,
                        0x00002d010300040000002d010500040000002d010100040000002701ffff0300 ,
                        0x00001e00040000002d010000040000002d010600040000002d01040005000000 ,
                        0x0102ffffff000500000009020000000007000000160439009b0213006e020700 ,
                        0x0000fc02000080206000ffff040000002d01020004000000f001060005000000 ,
                        0x0902ffffff00050000000102802060000400000004010d000400000002010200 ,
                        0x070000001b04340079022d00720205000000090200000000050000000102ffff ,
                        0xff000400000002010100040000002e0118000f000000320a34007c0205000000 ,
                        0x4769726c730008000300050003000700040000002e010000040000002d010300 ,
                        0x040000002d010500040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010200040000002d010400050000000102ffffff00 ,
                        0x0500000009020000000007000000160439009b0213006e02040000002d010300 ,
                        0x040000002d010500040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010200040000002d010400050000000102ffffff00 ,
                        0x050000000902000000000700000016045001a8020000000007000000fc020000 ,
                        0x000000000000040000002d01060004000000f0010200040000002d0103000400 ,
                        0x0000f0010400040000002701ffff050000000c025001a802030000001e000500 ,
                        0x00000102ffffff0005000000090200000000040000002701ffff050000000b02 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000000400 ,
                        0x00002701ffff0300000000007269616c3d00120092040c035028841200008a82 ,
                        0x77284776f48285000700710300000002000a0000000908080080050080e209c9 ,
                        0x07ac02020038005210040001021000331000008c00040001003d002610020003 ,
                        0x00531004000000090054100400000003005510060000000000000104000e0000 ,
                        0x00000000000006485f436f646504000c000000010000000004426f797304000d ,
                        0x0000000200000000054769726c7304000c0001000000000000044341524d0300 ,
                        0x0f0001000100000000000000000000494003000f000100020000000000000000 ,
                        0x0080514004000a000200000000000002434803000f0002000100000000000000 ,
                        0x000000444003000f00020002000000000000000000004c4004000c0003000000 ,
                        0x000000044348495303000f00030001000000000000000000c0524003000f0003 ,
                        0x0002000000000000000000405a4004000b000400000000000003434f4303000f ,
                        0x00040001000000000000000000003e4003000f00040002000000000000000000 ,
                        0x00454004000b0005000000000000034a504303000f0005000100000000000000 ,
                        0x000000344003000f00050002000000000000000000003c4004000b0006000000 ,
                        0x0000000353504303000f00060001000000000000000000004e4003000f000600 ,
                        0x0200000000000000000000554004000b00070000000000000353544d03000f00 ,
                        0x0700010000000000
                    End
                    RowSourceType ="Table/Query"
                    RowSource ="TRANSFORM Sum([HousePoints-Total-Sex-F].SumOfPoints) AS SumOfSumOfPoints SELECT "
                        "[HousePoints-Total-Sex-F].H_Code FROM [HousePoints-Total-Sex-F] GROUP BY [HouseP"
                        "oints-Total-Sex-F].H_Code PIVOT [HousePoints-Total-Sex-F].[HousePoints-Total-Sex"
                        "].[Sex Sub];"
                    Class ="MSGraph.Chart.5"
                    OLEClass ="Microsoft Graph 5.0"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =360
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

Option Explicit

Dim NumberToDisplay As Variant

' Generate HTML Variables and Constants
Dim sHTML As String, rHTML As String, PageNum As Integer, OldPg As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String, GenerateHTML As Boolean, aIndex As Integer
Dim ExportOleChart  As Boolean

Dim HTM() As HTMarrayType

Const ReportTitle = "Overall Results - By Gender"
Const repName = "sex" ' Keep to 4 letters or less (and unique from all other reports





Private Sub AddToArray(GrpName As Variant, GrpHead As Integer, s As String)

On Error Resume Next
    
    aIndex = aIndex + 1
    
    ReDim Preserve HTM(aIndex) As HTMarrayType
    HTM(aIndex).Pg = PageNum
    HTM(aIndex).GrpName = GrpName
    HTM(aIndex).GrpHead = GrpHead
    HTM(aIndex).row = s

End Sub

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    
    If GenerateHTML And Not Cancel And FormatCount = 1 Then
        
        DetailCount = DetailCount + 1
        
        If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
        
        rHTML = ""
        Call RowStart(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Place)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_NAme)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_Code)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!SumOfPoints)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        Call AddToArray(Me!Age, rDetail, rHTML)
    End If

    '*** HTML Generation Code End ***


End Sub



Private Sub Group1()

On Error Resume Next

    '*** HTML Generation Code Start ***

    rHTML = ""
    ' *** Create Group Title
    Call RowStart(rHTML)

    Call CellStart(rHTML, "", "", "10%", cWhite, 5)
    rHTML = rHTML & Heading(3, "GENDER: " & Me!Age, 3)
    Call CellEnd(rHTML)
    
    Call RowEnd(rHTML)
    
    ' *** Create general record header ***
    Call RowStart(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "PLACE")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "45%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM NAME")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "30%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM CODE")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "15%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TOTAL")
    Call CellEnd(rHTML)

    Call RowEnd(rHTML)

    Call AddToArray(Me!Age, rGroupHeader, rHTML)

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        rHTML = ""
        Call RowStart(rHTML)
    
        Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        Call AddToArray(Me!Age, rGroupFooter, rHTML)

    End If


End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        Call Group1
    End If

End Sub


Private Sub PageFooter2_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
                
        rHTML = ""
        Call TableEnd(rHTML)
        Call AddToArray(Me!Age, rPageFooter, rHTML)
        
    End If


End Sub


Private Sub PageHeader0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        'NewPage = True
        
        'DetailCount = 0
        PageNum = PageNum + 1
        rHTML = ""
        
        If PageNum > 1 Then
            PrevPage = Link(repName & PageNum - 1 & ".htm", "Previous Page")
        Else
            PrevPage = ""
        End If
        NextPage = Link(repName & PageNum + 1 & ".htm", "Next Page")
        
        Call TableStart(rHTML, "95%", "", "", "", 0)
        
        Call AddToArray(Me!Age, rPageHeader, rHTML)

    End If


End Sub

Private Sub Report_Close()
On Error GoTo Report_Close_Err
  Dim HTMLFileLocation As String, FileLocation As String
  If ExportOleChart Then
    Dim oleGraph As Object
    HTMLFileLocation = DLookup("[HTMLlocation]", "MiscHTML")
    FileLocation = HTMLFileLocation & "\sex.jpg"
    Set oleGraph = Me.oleChart.Object
    
    oleGraph.export fileName:=FileLocation
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

    ' *** HTML Creation Code ***
    ExportOleChart = GlobalGenerateHTML
    GenerateHTML = GlobalGenerateHTML
    
    If GenerateHTML Then
        aIndex = 0
        PleaseWaitMsg = "Preparing HTML for """ & ReportTitle & """.  Please wait..."
        DoCmd.RunMacro "ShowPleaseWait"
    End If
    
    PageNum = 0
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
    ' ***************************

End Sub

Private Sub ReportFooter1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

Dim gHeader As Integer, OldPg As Integer, OldGroupName As String, i As Integer
Dim NewPg As Integer

    If GenerateHTML Then
        Dim eHTML As String, AlleHTML As String, sEvents   As String

        GenerateHTML = False
        
        rHTML = ""
        Call TableEnd(rHTML)
    
        'Debug.Print "RF - FormatCount="; FormatCount; " Page="; PageNum;  Me!Age
        Call AddToArray(Me!Age, False, rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "5%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "95%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "AGE(S)")
        Call CellEnd(sHTML)
        
        Call RowEnd(sHTML)

    
        DoCmd.RunMacro "ClosePleaseWait"
        
        OldPg = HTM(aIndex).Pg
        gHeader = False
        OldGroupName = HTM(aIndex).GrpName
        
        ' Initialise variables to create Summary Page
        sEvents = OldGroupName
        eHTML = ""
        AlleHTML = ""
        
        rHTML = ""
        
        For i = aIndex To 1 Step -1
            
            NewPg = HTM(i).Pg
            If HTM(i).GrpHead = rPageHeader Then
                
                ' *** Create HTML Page
                rHTML = HTM(i).row & rHTML
                If OldPg > 1 Then
                    PrevPage = Link(repName & OldPg - 1 & ".htm", "Previous Page")
                Else
                    PrevPage = ""
                End If
                If OldPg < HTM(aIndex).Pg Then
                    NextPage = Link(repName & OldPg + 1 & ".htm", "Next Page")
                Else
                    NextPage = ""
                End If
                rHTML = rHTML & " <p align=""center""><img border=""0"" src=""sex.jpg"" </p>"
                Call CreateHTMLfile(repName & OldPg & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & OldPg, ReportHead)
                rHTML = ""
                
                ' *** Create summary record ***
                If OldPg Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
                
                Call RowStart(eHTML)
    
                Call CellStart(eHTML, "Center", "", "5%", BGcolor, 1)
                eHTML = eHTML & LinkStart(repName & OldPg & ".htm")
                Call Text(eHTML, "", "", Str(OldPg))
                eHTML = eHTML & LinkEnd()
                Call CellEnd(eHTML)
    
                Call CellStart(eHTML, "", "", "95%", BGcolor, 1)
                Call Text(eHTML, "", "", sEvents)
                Call CellEnd(eHTML)
                
                Call RowEnd(eHTML)
        
                AlleHTML = eHTML & AlleHTML
                eHTML = ""
                sEvents = ""

            End If
            
            If (HTM(i).GrpHead = rGroupHeader) And Not gHeader Then
                gHeader = True
                rHTML = HTM(i).row & rHTML
            End If
            
            If OldGroupName = HTM(i).GrpName And Not gHeader Then
                rHTML = HTM(i).row & rHTML
            
            ElseIf (OldGroupName <> HTM(i).GrpName) And (HTM(i).GrpHead <> rPageFooter) Then
                Dim SpacedEvent As String

                SpacedEvent = HTM(i).GrpName
                Call SpaceIndent(SpacedEvent, 5)
                sEvents = SpacedEvent & " " & sEvents
                rHTML = HTM(i).row & rHTML
                gHeader = False
            End If
            
            If (HTM(i).GrpHead = rGroupHeader) And (OldGroupName <> HTM(i).GrpName) Then
                gHeader = True
                rHTML = HTM(i).row & rHTML
            End If
            
            
            'Debug.Print HTM(i).Pg, HTM(i).GrpName, HTM(i).GrpHead', HTM(i).row

            ' Ignore PageFooter groupType.  I hope it is not needed ever
            If (HTM(i).GrpHead <> rPageFooter) Then
                OldGroupName = HTM(i).GrpName
            End If
            OldPg = NewPg
        Next

        ' * Generate Summary Page file
        sHTML = sHTML & AlleHTML
        Call TableEnd(sHTML)
        sHTML = sHTML & " <p align=""center""><img border=""0"" src=""sex.jpg"" </p>"
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead)


    End If

End Sub
