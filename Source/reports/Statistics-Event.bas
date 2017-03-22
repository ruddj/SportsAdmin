Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =117
    Left =1905
    Top =240
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x877fb034ecdae140
    End
    RecordSource ="HousePoints-Total-Event-F"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000045020000d002000000000000a02900006301000001000000 ,
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
            ControlSource ="ET_Des"
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
                    Caption ="Overall Statistical Summary - by Event"
                    FontName ="times New Roman"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1045
            OnFormat ="[Event Procedure]"
            OnRetreat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =1267
                    Top =685
                    Width =1515
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
                    Left =5782
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
                    Left =60
                    Top =120
                    Width =1215
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text114"
                    Caption ="EVENT:"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =1303
                    Top =170
                    Width =6756
                    Height =330
                    FontSize =11
                    FontWeight =700
                    Name ="ET_Des"
                    ControlSource ="ET_Des"

                End
            End
        End
        Begin Section
            Height =355
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =1267
                    Width =4341
                    Height =330
                    FontSize =10
                    FontWeight =700
                    Name ="H_NAme"
                    ControlSource ="H_NAme"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5785
                    Width =1536
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="H_Code"
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
                    Name ="SumOfPoints"
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
            Height =170
            OnFormat ="[Event Procedure]"
            OnRetreat ="[Event Procedure]"
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =5775
            OnFormat ="[Event Procedure]"
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =5385
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =2
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =5385
                    Width =9426
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
                    Top =5385
                    Width =10596
                    Name ="Line87"
                End
                Begin Chart
                    ColumnHeads = NotDefault
                    Locked = NotDefault
                    SizeMode =3
                    RowSourceTypeInt =2
                    Left =160
                    Top =170
                    Width =10319
                    Height =5113
                    TabIndex =1
                    Name ="oleChart"
                    OleData = Begin
                        0x007c0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000000000000000000000100000 ,
                        0x0f00000001000000feffffff0000000001000000ffffffffffffffffffffffff ,
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
                        0x0000000016000500ffffffffffffffff050000000308020000000000c0000000 ,
                        0x00000046000000000000000000000000e040883c1facc00110000000c00d0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000064000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000003600000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000006000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000030000009d090000 ,
                        0x000000000e000000fdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff3c000000feffffff ,
                        0xfeffffff15000000120000001300000019000000ffffffff1600000017000000 ,
                        0x180000000d0000001a0000001b00000021000000ffffffffffffffffffffffff ,
                        0xffffffffffffffff22000000230000002400000025000000feffffff27000000 ,
                        0x28000000290000002a0000002b00000030000000ffffffffffffffffffffffff ,
                        0xffffffff31000000feffffff3300000034000000350000003600000037000000 ,
                        0x38000000390000003a0000003b00000026000000feffffffffffffffffffffff ,
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
                        0x00000046000000000000000000000000c0357f6c58b9bf010600000080170000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000062000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000003600000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000006000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000030000009d090000 ,
                        0x00000000ffffffffffffffff04000000fdfffffffefffffffeffffff15000000 ,
                        0x08000000090000000a0000000b000000feffffff070000000c000000ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff1c0000001600000017000000 ,
                        0x180000000d000000ffffffffffffffffffffffff1d0000001e0000001f000000 ,
                        0x200000002c000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff2d0000002e0000002f000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
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
                        0x0000000000000000000000000000000000000000000000001400000038130000 ,
                        0x0000000057006f0072006b0062006f006f006b00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001200020101000000ffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002a000000a30c0000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x2800000029000000feffffff2b0000002c0000002d0000002e0000002f000000 ,
                        0x30000000310000003200000033000000340000003500000037000000feffffff ,
                        0x38000000390000003a0000003b0000003c0000003d0000003e0000003f000000 ,
                        0x4000000041000000420000004300000044000000450000004600000047000000 ,
                        0x48000000490000004a0000004b0000004c0000004d0000004e0000004f000000 ,
                        0x5000000051000000520000005300000054000000550000005600000057000000 ,
                        0x58000000590000005a0000005b0000005c0000005d000000feffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0308020000000000c000000000000046 ,
                        0x130000004d6963726f736f667420477261706820393700070000004742696666 ,
                        0x3500100000004d5347726170682e43686172742e3800f439b271000000000000 ,
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
                        0x23302e3061006c0031001a00c3000100ff7fbc02000000000000050141007200 ,
                        0x690061006c0031001a00a0000100ff7fbc020000000200000501410072006900 ,
                        0x61006c0031001a00c8000100ff7fbc0200000000000005014100720069006100 ,
                        0x6c003d0012003804d002c828901500002c2323302e30305f85000800c4030000 ,
                        0x000200000a0000000908100080060080b80dcc074100000006000000ac020200 ,
                        0x38009200e200380000000000ffffff00ff00000000ff00000000ff00ffff0000 ,
                        0xff00ff0000ffff00800000000080000000008000808000008000800000808000 ,
                        0xc0c0c000808080008080ff0080206000ffffc000a0e0e00060008000ff808000 ,
                        0x0080c000c0c0ff0000008000ff00ff00ffff000000ffff008000800080000000 ,
                        0x008080000000ff0000ccff0069ffff00ccffcc00ffff9900a6caf000cc9ccc00 ,
                        0xcc99ff00e3e3e3003366ff0033cccc0033993300999933009966330099666600 ,
                        0x66669900969696003333cc003366660000330000333300006633000099336600 ,
                        0x33339900424242005c100e00030000000000ffffff0000000000521004000102 ,
                        0x1000331000008c00040001003d00261002000500531004000000040054100400 ,
                        0x000006005510060000000000000104000f00000000000000000600485f436f64 ,
                        0x6504001400000001000000000b003130306d20537072696e7404001400000002 ,
                        0x000000000b003230306d20537072696e7404000d000000030000000004003830 ,
                        0x306d0400120000000400000000090048696768204a756d700400120000000500 ,
                        0x00000009004c6f6e67204a756d7004000c00010000000000000300434f430300 ,
                        0x0f0001000100000000000000000000404003000f000100020000000000000000 ,
                        0x0000204004000c000200000000000003004a504303000f000200010000000000 ,
                        0x00000000003c4003000f00020002000000000000000000001c4004000d000300 ,
                        0x000000000004005452494e03000f000300010000000000000000000032400300 ,
                        0x0f00030002000000000000000000001840571001000159100800b4001d01eb32 ,
                        0xfd203d000a0000000000f627000f00003e000e00010101000101000100010000 ,
                        0x00005810020000001d0011000301000100000001000100010001000100341000 ,
                        0x0001100200000002101000000000000000000000000102d0bffc0033100000a0 ,
                        0x0004000100010064100800000001000000010003100c00030001000300030001 ,
                        0x000000331000005110080000010200000001000d101a0000000b013100300030 ,
                        0x006d00200053007000720069006e007400511008000101020000000100511008 ,
                        0x00020102000000000051100800030102000000000006100800ffff0000000000 ,
                        0x00331000005f1002000000341000004510020000003410000003100c00030001 ,
                        0x000300030001000000331000005110080000010200000002000d101a0000000b ,
                        0x013200300030006d00200053007000720069006e007400511008000101020000 ,
                        0x00020051100800020102000000000051100800030102000000000006100800ff ,
                        0xff010001000000331000005f1002000000341000004510020000003410000003 ,
                        0x100c00030001000300030001000000331000005110080000010200000003000d ,
                        0x100c00000004013800300030006d005110080001010200000003005110080002 ,
                        0x0102000000000051100800030102000000000006100800ffff02000200000033 ,
                        0x1000005f1002000000341000004510020000003410000003100c000300010003 ,
                        0x00030001000000331000005110080000010200000004000d1016000000090148 ,
                        0x0069006700680020004a0075006d007000511008000101020000000400511008 ,
                        0x00020102000000000051100800030102000000000006100800ffff0300030000 ,
                        0x00331000005f1002000000341000004510020000003410000003100c00030001 ,
                        0x000300030001000000331000005110080000010200000005000d101600000009 ,
                        0x014c006f006e00670020004a0075006d00700051100800010102000000050051 ,
                        0x100800020102000000000051100800030102000000000006100800ffff040004 ,
                        0x000000331000005f100200000034100000451002000000341000004410040019 ,
                        0x000000241002000200251020000202010000000000e2ffffffc3ffffff000000 ,
                        0x0000000000b1004d0020100000331000004f1014000200020000000000000000 ,
                        0x0000000000000000002610020003005110080000010200000000003410000024 ,
                        0x1002000300251020000202010000000000e2ffffffc3ffffff00000000000000 ,
                        0x00b1004d0020100000331000004f101400020002000000000000000000000000 ,
                        0x0000000000261002000300511008000001020000000000341000004610020001 ,
                        0x0041101200000083000000870100008b0b0000230c0000331000004f10140002 ,
                        0x00020000000000090100000d0c0000ee0d00001d101200000000000000000000 ,
                        0x0000000000000000003310000020100800010001000100010062101200000000 ,
                        0x000100000001000000000000006f001e101e0002000301000000000000000000 ,
                        0x000000000000000000000023004d000000341000001d10120001000000000000 ,
                        0x0000000000000000000000331000001f102a0000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000001f011e101e ,
                        0x0002000301000000000000000000000000000000000000000023004d00000034 ,
                        0x1000001410140000000000000000000000000000000000000000003310000017 ,
                        0x10060000009600000022100a0000000000000000000f00151014002c0c000000 ,
                        0x0000003f030000a00f000003010200331000004f101400050001002b0c0000f4 ,
                        0xffffff69000000f5000000251020000202010000000000e2ffffffc3ffffff00 ,
                        0x00000000000000b1004d0000000000331000004f101400020002000000000000 ,
                        0x0000000000000000000000261002000300511008000001020000000000341000 ,
                        0x0032100400000002003310000007100c00000000000500000008004d000a1010 ,
                        0x00ffffff0000000000000000004e004d00341000003410000034100000341000 ,
                        0x002510200002020100000000007206000056000000bc0200000101000081404d ,
                        0x0030100000331000004f10140002000200000000000000000076000000150000 ,
                        0x005110080000010200000000000d102e000000150154006f00740061006c0020 ,
                        0x0050006f0069006e007400730020006200790020004500760065006e00740027 ,
                        0x100600010000000000341000003410000000020e000000000003000000000005 ,
                        0x0000000a00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000005f2d3b5f2d405f2d1e041c00a400170000222422232c2323305f293b ,
                        0x5c28222422232c2323305c291e042100a5001c0000222422232c2323305f293b ,
                        0x5b5265645d5c28222422232c2323305c291e042200a6001d0000222422232c23 ,
                        0x23302e30305f293b5c28222422232c2323302e30305c291e042700a700220000 ,
                        0x222422232c2323302e30305f293b5b5265645d5c28222422232c2323302e3030 ,
                        0x5c291e043700a8003200005f282224222a20232c2323305f293b5f282224222a ,
                        0x205c28232c2323305c293b5f282224222a20222d225f293b5f28405f291e042e ,
                        0x00a9002900005f282a20232c2323305f293b5f282a205c28232c2323305c293b ,
                        0x5f282a20222d225f293b5f28405f291e043f00aa003a00005f282224222a2023 ,
                        0x2c2323302e30305f293b5f282224222a205c28232c2323302e30305c293b5f28 ,
                        0x2224222a20222d223f3f5f293b5f28405f291e043600ab003100005f282a2023 ,
                        0x2c2323302e30305f293b5f282a205c28232c2323302e30305c293b5f282a2022 ,
                        0x2d223f3f01000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000005f293b5f28405f2931001a00a0000100ff7fbc020000000000000501 ,
                        0x41007200690061006c0031001a00a0000100ff7fbc0200000000000005014100 ,
                        0x720069000000000000000000007c000000190000005110080000010200000000 ,
                        0x000d101800000015546f74616c20506f696e7473206279204576656e74271006 ,
                        0x00010000000000341000003410000000000a00000006000000050000000a0000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000908100080060500b80dcc07410000000600000042000200b0042200 ,
                        0x020000001e0418000500130000222422232c2323303b5c2d222422232c232330 ,
                        0x1e041d000600180000222422232c2323303b5b5265645d5c2d222422232c2323 ,
                        0x301e041e000700190000222422232c2323302e30303b5c2d222422232c232330 ,
                        0x2e30301e04230008001e0000222422232c2323302e30303b5b5265645d5c2d22 ,
                        0x2422232c2323302e30301e0435002a003000005f2d2224222a20232c2323305f ,
                        0x2d3b5c2d2224222a20232c2323305f2d3b5f2d2224222a20222d225f2d3b5f2d ,
                        0x405f2d1e042c0029002700005f2d2a20232c2323305f2d3b5c2d2a20232c2323 ,
                        0x305f2d3b5f2d2a20222d225f2d3b5f2d405f2d1e043d002c003800005f2d2224 ,
                        0x222a20232c2323302e30305f2d3b5c2d2224222a20232c2323302e30305f2d3b ,
                        0x5f2d2224222a20222d223f3f5f2d3b5f2d405f2d1e0434002b002f00005f2d2a ,
                        0x20232c2323302e30305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a2022 ,
                        0x2d223f3f03004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000180002010300000004000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000003200000042230000 ,
                        0x0000000057006f0072006b0062006f006f006b00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001200020101000000ffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000110000003b140000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x2800000029000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffeffffff ,
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
                        0x000000000100feff030a0000ffffffff0308020000000000c000000000000046 ,
                        0x150000004d6963726f736f667420477261706820323030300007000000474269 ,
                        0x66663500100000004d5347726170682e43686172742e3800f439b27100000000 ,
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
                        0x23302e300908100080060500af18cd07414000000601000042000200b0042200 ,
                        0x020000001e0418000500130000222422232c2323303b5c2d222422232c232330 ,
                        0x1e041d000600180000222422232c2323303b5b5265645d5c2d222422232c2323 ,
                        0x301e041e000700190000222422232c2323302e30303b5c2d222422232c232330 ,
                        0x2e30301e04230008001e0000222422232c2323302e30303b5b5265645d5c2d22 ,
                        0x2422232c2323302e30301e0435002a003000005f2d2224222a20232c2323305f ,
                        0x2d3b5c2d2224222a20232c2323305f2d3b5f2d2224222a20222d225f2d3b5f2d ,
                        0x405f2d1e042c0029002700005f2d2a20232c2323305f2d3b5c2d2a20232c2323 ,
                        0x305f2d3b5f2d2a20222d225f2d3b5f2d405f2d1e043d002c003800005f2d2224 ,
                        0x222a20232c2323302e30305f2d3b5c2d2224222a20232c2323302e30305f2d3b ,
                        0x5f2d2224222a20222d223f3f5f2d3b5f2d405f2d1e0434002b002f00005f2d2a ,
                        0x20232c2323302e30305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a2022 ,
                        0x2d223f3f5f2d3b5f2d405f2d1e041c00a400170000222422232c2323305f293b ,
                        0x5c28222422232c2323305c291e042100a5001c0000222422232c2323305f293b ,
                        0x5b5265645d5c28222422232c2323305c291e042200a6001d0000222422232c23 ,
                        0x23302e30305f293b5c28222422232c2323302e30305c291e042700a700220000 ,
                        0x222422232c2323302e30305f293b5b5265645d5c28222422232c2323302e3030 ,
                        0x5c291e043700a8003200005f282224222a20232c2323305f293b5f282224222a ,
                        0x205c28232c2323305c293b5f282224222a20222d225f293b5f28405f291e042e ,
                        0x00a9002900005f282a20232c2323305f293b5f282a205c28232c2323305c293b ,
                        0x5f282a20222d225f293b5f28405f291e043f00aa003a00005f282224222a2023 ,
                        0x2c2323302e30305f293b5f282224222a205c28232c2323302e30305c293b5f28 ,
                        0x2224222a20222d223f3f5f293b5f28405f291e043600ab003100005f282a2023 ,
                        0x2c2323302e30305f293b5f282a205c28232c2323302e30305c293b5f282a2022 ,
                        0x2d223f3f5f293b5f28405f2931001a00a0000100ff7fbc020000000000000501 ,
                        0x41007200690061006c0031001a00a0000100ff7fbc0200000000000005014100 ,
                        0x7200690061006c0031001a00c3000100ff7fbc02000000000000050141007200 ,
                        0x690061006c0031001a00a0000100ff7fbc020000000200000501410072006900 ,
                        0x61006c0031001a00c8000100ff7fbc0200000000000005014100720069006100 ,
                        0x6c003d0012003804d002c828901500002c2323302e30305f85000800c4030000 ,
                        0x000200000a0000000908100080060080af18cd074140000006010000ac020200 ,
                        0x38009200e200380000000000ffffff00ff00000000ff00000000ff00ffff0000 ,
                        0xff00ff0000ffff00800000000080000000008000808000008000800000808000 ,
                        0xc0c0c000808080008080ff0080206000ffffc000a0e0e00060008000ff808000 ,
                        0x0080c000c0c0ff0000008000ff00ff00ffff000000ffff008000800080000000 ,
                        0x008080000000ff0000ccff0069ffff00ccffcc00ffff9900a6caf000cc9ccc00 ,
                        0xcc99ff00e3e3e3003366ff0033cccc0033993300999933009966330099666600 ,
                        0x66669900969696003333cc003366660000330000333300006633000099336600 ,
                        0x33339900424242005c100e00030000000000ffffff0000000000521004000102 ,
                        0x1000331000008c00040001003d00261002000500531004000000050054100400 ,
                        0x00000d005510060000000000000104000f00000000000000000600485f436f64 ,
                        0x6504001800000001000000000f003130306d204261636b7374726f6b6504001a ,
                        0x000000020000000011003130306d204272656173747374726f6b650400170000 ,
                        0x0003000000000e003130306d20427574746572666c7904001700000004000000 ,
                        0x000e003130306d20467265657374796c6504001f000000050000000016003230 ,
                        0x306d20496e646976696475616c204d65646c657904001b000000060000000012 ,
                        0x00347835306d204d65646c65792052656c617904001400000007000000000b00 ,
                        0x347835306d2052656c617904001700000008000000000e0035306d204261636b ,
                        0x7374726fffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xb2460000d4220000fe1200000100090000037f09000007002700000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c025101ac02030000001e00040000002701ffff030000001e00 ,
                        0x040000002701ffff050000000b0200000000030000001e00050000000102ffff ,
                        0xff000500000009020000000010000000fb02f5ff000000000000bc0200000000 ,
                        0x00000022417269616c000000040000002d010000040000002d01000004000000 ,
                        0x2d01000010000000fb021000070000000000bc02000000000102022253797374 ,
                        0x656d0000040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000050000000102ffffff00050000000902000000000700000016045101ac02 ,
                        0x00000000040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000050000000102ffffff00050000000902000000000700000016045101ac02 ,
                        0x000000000700000015044c019e0205001202040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x000000000700000016045101ac0200000000040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x00000000305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a20222d223f3f ,
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
                        0xff7fbc0200000000004505417269616c31001400c8000100ff7fbc0200000000 ,
                        0x004505417269616c31001400a0000100ff7fbc0200000002004505417269616c ,
                        0x31001400c8000100ff7fbc0200000000004505417269616c3d001200a005c003 ,
                        0xab27ae15000000003e200000000085000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x003d002610020003005310040000000400541004000000060055100600000000 ,
                        0x00000104000e000000000000000006485f436f64650400130000000100000000 ,
                        0x0b3130306d20537072696e7404001300000002000000000b3230306d20537072 ,
                        0x696e7404000c0000000300000000043830306d04001100000004000000000948 ,
                        0x696768204a756d700400110000000500000000094c6f6e67204a756d7004000b ,
                        0x000100000000000003434f4303000f0001000100000000000000000000404003 ,
                        0x000f0001000200000000000000000000204004000b0002000000000000034a50 ,
                        0x4303000f00020001000000000000000000003c4003000f000200020000000000 ,
                        0x00000000001c4004000c0003000000000000045452494e03000f000300010000 ,
                        0x0000000000000000324003000f00030002000000000000000000001840571001 ,
                        0x000159100800b4002c01eb32681f3d000a0000000000f627000f00003e000e00 ,
                        0x01010100010100010001000000005810020000001d0011000301000100000001 ,
                        0x0001000100010001003410000001100200000002101000000000000000000000 ,
                        0xc0f501d03f010133100000a00004000100010003100800030001000300030033 ,
                        0x1000005110080000010200000001000d100e0000000b3130306d20537072696e ,
                        0x7451100800010102000000010051100800020102000000000006100800ffff00 ,
                        0x0000000000451002000000341000000310080003000100030003003310000051 ,
                        0x10080000010200000002000d100e0000000b3230306d20537072696e74511008 ,
                        0x00010102000000020051100800020102000000000006100800ffff0100010000 ,
                        0x0045100200000034100000031008000300010003000300331000005110080000 ,
                        0x010200000003000d1007000000043830306d5110080001010200000003005110 ,
                        0x0800020102000000000006100800ffff02000200000045100200000034100000 ,
                        0x031008000300010003000300331000005110080000010200000004000d100c00 ,
                        0x00000948696768204a756d705110080001010200000004005110080002010200 ,
                        0x0000000006100800ffff03000300000045100200000034100000031008000300 ,
                        0x010003000300331000005110080000010200000005000d100c000000094c6f6e ,
                        0x67204a756d705110080001010200000005005110080002010200000000000610 ,
                        0x0800ffff04000400000045100200000034100000441003000900004610020001 ,
                        0x00411012000000860000008c010000740b0000270c0000331000004f10140002 ,
                        0x000200000000001a010000f90b0000e00d00001d101200000000000000000000 ,
                        0x000000000000000000331000001e101a00020003010000000000000000000000 ,
                        0x0000000000000000002300341000001d10120001000000000000000000000000 ,
                        0x0000000000331000001f102a0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000001f011e101a000200030100 ,
                        0x0000000000000000000000000000000000000023003410000014101400000000 ,
                        0x0000000000000000000000000000000000331000001710060000009600000022 ,
                        0x100a0000000000000000000f0015101400180c00003000000052030000ad0500 ,
                        0x0003010200331000004f10140005000100180c000031000000690000005b0000 ,
                        0x0025101a000202010000000000e2ffffffc4ffffff0000000000000000b10033 ,
                        0x1000004f10140002000200000000000000000000000000000000005110080000 ,
                        0x010200000000003410000032100400000002003310000007100a000000000000 ,
                        0x00000009000a100c00ffffff0000000000000000003410000034100000241002 ,
                        0x00020025101a000202010000000000e2ffffffc4ffffff0000000000000000b1 ,
                        0x00331000004f1014000200020000000000000000000000000000000000261002 ,
                        0x00020051100800000102000000000034100000341000003410000025101a0002 ,
                        0x020100000000005506000054000000f10200002c0100008100331000004f1014 ,
                        0x000200026b650400190000000900000000100035306d20427265617374737472 ,
                        0x6f6b650400160000000a000000000d0035306d20427574746572666c79040016 ,
                        0x0000000b000000000d0035306d20467265657374796c650400130000000c0000 ,
                        0x00000a004b69636b20426f61726404000e000100000000000005004153484552 ,
                        0x03000f0001000100000000000000000000514003000f00010002000000000000 ,
                        0x00000080424003000f0001000300000000000000000080414003000f00010004 ,
                        0x000000000000000000c0504003000f00010005000000000000000000004f4003 ,
                        0x000f00010008000000000000000000c0664003000f0001000900000000000000 ,
                        0x000080654003000f0001000a000000000000000000804c4003000f0001000b00 ,
                        0x0000000000000000c06c4003000f0001000c0000000000000000008040400400 ,
                        0x10000200000000000007004550485241494d03000f0002000100000000000000 ,
                        0x000000404003000f0002000200000000000000000000354003000f0002000300 ,
                        0x000000000000000000394003000f000200040000000000000000000022400300 ,
                        0x0f0002000500000000000000000080434003000f000200080000000000000000 ,
                        0x0060634003000f00020009000000000000000000a0674003000f0002000a0000 ,
                        0x00000000000000804a4003000f0002000b000000000000000000606d4003000f ,
                        0x0002000c00000000000000000080424004000e000300000000000005004a5544 ,
                        0x414803000f0003000100000000000000000080414003000f0003000200000000 ,
                        0x000000000000304003000f00030003000000000000000000002c4003000f0003 ,
                        0x000400000000000000000080434003000f000300050000000000000000000049 ,
                        0x4003000f00030008000000000000000000005f4003000f000300090000000000 ,
                        0x00000000c0644003000f0003000a000000000000000000c0584003000f000300 ,
                        0x0b000000000000000000a06a4003000f0003000c000000000000000000004140 ,
                        0x04000d000400000000000004004c45564903000f000400010000000000000000 ,
                        0x0080404003000f0004000200000000000000000000324003000f000400030000 ,
                        0x00000000000000003d4003000f0004000400000000000000000000264003000f ,
                        0x0004000500000000000000000000424003000f00040008000000000000000000 ,
                        0x00644003000f0004000900000000000000000080664003000f0004000a000000 ,
                        0x00000000000000504003000f0004000b000000000000000000a0684003000f00 ,
                        0x04000c000000000000000000003e40571001000159100800b4001d01eb32fd20 ,
                        0x3d000a0000000000f627000f00003e000e000101010001010001000100000000 ,
                        0x5810020000001d00110003010001000000010001000100010001003410000001 ,
                        0x100200000002101000000000000000000000000102d0bffc0033100000a00004 ,
                        0x000100010064100800000001000000010003100c000300010004000400010000 ,
                        0x00331000005110080000010200000001000d10220000000f013100300030006d ,
                        0x0020004200610063006b007300740072006f006b006500511008000101020000 ,
                        0x00010051100800020102000000000051100800030102000000000006100800ff ,
                        0xff000000000000331000005f1002000000341000004510020000003410000003 ,
                        0x100c00030001000400040001000000331000005110080000010200000002000d ,
                        0x102600000011013100300030006d002000420072006500610073007400730074 ,
                        0x0072006f006b0065005110080001010200000002005110080002010200000000 ,
                        0x0051100800030102000000000006100800ffff010001000000331000005f1002 ,
                        0x000000341000004510020000003410000003100c000300010004000400010000 ,
                        0x00331000005110080000010200000003000d10200000000e013100300030006d ,
                        0x00200042007500740074006500720066006c0079005110080001010200000003 ,
                        0x0051100800020102000000000051100800030102000000000006100800ffff02 ,
                        0x0002000000331000005f1002000000341000004510020000003410000003100c ,
                        0x00030001000400040001000000331000005110080000010200000004000d1020 ,
                        0x0000000e0700000016044d01a80204000400040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x0000000007000000160425010f0224001a00040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x000000000700000016042401100222001a0009000000fa020000010000000000 ,
                        0x00002200040000002d01020007000000fc0200008080ff008020040000002d01 ,
                        0x0300050000000902ffffff000500000001028080ff000400000004010d000400 ,
                        0x0000020102000e000000240305002e003b0048003b00480023012e0023012e00 ,
                        0x3b000e00000024030500d4005800ee005800ee002301d4002301d40058000e00 ,
                        0x0000240305007a01a0009401a000940123017a0123017a01a00007000000fc02 ,
                        0x0000802060008020040000002d01040004000000f00103000500000001028020 ,
                        0x60000e000000240305004800e9006100e90061002301480023014800e9000e00 ,
                        0x000024030500ee00f0000701f00007012301ee002301ee00f0000e0000002403 ,
                        0x05009401f700ad01f700ad012301940123019401f70009000000fa0200000000 ,
                        0x0000000000002200040000002d01030007000000fc020000ffffff0000000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010200050000000102802060000500 ,
                        0x00000902ffffff0007000000160425010f0224001a00040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010200050000000102802060000500 ,
                        0x00000902ffffff000700000016045101ac020000000009000000fa0200000100 ,
                        0x0000000000002200040000002d01060004000000f00102000500000014022500 ,
                        0x1b00050000000102ffffff000400000004010d00040000000201010005000000 ,
                        0x130223011b000500000014022301180005000000130223011b00050000001402 ,
                        0xff001800050000001302ff001b00050000001402da001800050000001302da00 ,
                        0x1b00050000001402b6001800050000001302b6001b0005000000140292001800 ,
                        0x05000000130292001b000500000014026e0018000500000013026e001b000500 ,
                        0x000014024900180005000000130249001b000500000014022500180005000000 ,
                        0x130225001b0005000000140223011b0005000000130223010d02050000001402 ,
                        0x26011b0005000000130223011b000500000014022601c1000500000013022301 ,
                        0xc100050000001402260167010500000013022301670105000000140226010d02 ,
                        0x05000000130223010d02040000002d010000040000002d010000040000002d01 ,
                        0x0300040000002d010500040000002d010100040000002701ffff030000001e00 ,
                        0x040000002d010000040000002d010400040000002d010600050000000102ffff ,
                        0xff00050000000902ffffff00070000001604220092010b001a01050000000902 ,
                        0x000000000400000004010d00040000000201010027000000320a0f001d011500 ,
                        0x0000546f74616c20506f696e7473206279204576656e74000700070004000600 ,
                        0x0300030007000700030007000400070003000700060003000600060007000700 ,
                        0x0400040000002d010000040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010400 ,
                        0x040000002d010600050000000102ffffff000500000009020000000007000000 ,
                        0x16045101ac0200000000040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010400 ,
                        0x040000002d010600050000000102ffffff000500000009020000000007000000 ,
                        0x16045101ac0200000000040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010400 ,
                        0x040000002d010600050000000102ffffff000500000009020000000007000000 ,
                        0x16045201ad02ffffffff0400000004010d00040000000201010009000000320a ,
                        0x1c010d00010000003000060009000000320af8000d0001000000350006000a00 ,
                        0x0000320ad3000700020000003130060006000a000000320aaf00070002000000 ,
                        0x3135060006000a000000320a8b000700020000003230060006000a000000320a ,
                        0x67000700020000003235060006000a000000320a420007000200000033300600 ,
                        0x06000a000000320a1e00070002000000333506000600040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016045101ac0200000000040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016045201ad02ffffffff0400000004010d000400 ,
                        0x0000020101000c000000320a2c01620003000000434f43000800080008000c00 ,
                        0x0000320a2c010a01030000004a5043000600070008000d000000320a2c01ae01 ,
                        0x040000005452494e0700070003000700040000002d010300040000002d010500 ,
                        0x040000002d010100040000002701ffff030000001e00040000002d0100000400 ,
                        0x00002d010400040000002d010600050000000102ffffff000500000009020000 ,
                        0x00000700000016045101ac020000000009000000fa0205000000000000000000 ,
                        0x2200040000002d01020004000000f0010600040000002d010300040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010400040000002d010200050000000102ffffff00050000000902 ,
                        0x000000000700000016044d019f0204001102040000002d010300040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010400040000002d010200050000000102ffffff00050000000902 ,
                        0x000000000700000016044d019f020400110209000000fa020000010000000000 ,
                        0x00002200040000002d01060004000000f001020007000000fc0200008080ff00 ,
                        0x8020040000002d01020004000000f0010400050000000902ffffff0005000000 ,
                        0x01028080ff000400000004010d000400000002010200070000001b042a003b02 ,
                        0x2300340205000000090200000000050000000102ffffff000400000002010100 ,
                        0x040000002e01180018000000320a2a003e020b0000003130306d20537072696e ,
                        0x74000600060006000b000300070007000500030007000400040000002e010000 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010200040000002d0106000500 ,
                        0x00000102013100300030006d00200046007200650065007300740079006c0065 ,
                        0x0051100800010102000000040051100800020102000000000051100800030102 ,
                        0x000000000006100800ffff030003000000331000005f10020000003410000045 ,
                        0x10020000003410000003100c0003000100040004000100000033100000511008 ,
                        0x0000010200000005000d103000000016013200300030006d00200049006e0064 ,
                        0x006900760069006400750061006c0020004d00650064006c0065007900511008 ,
                        0x0001010200000005005110080002010200000000005110080003010200000000 ,
                        0x0006100800ffff040004000000331000005f1002000000341000004510020000 ,
                        0x003410000003100c000300010004000400010000003310000051100800000102 ,
                        0x00000006000d1028000000120134007800350030006d0020004d00650064006c ,
                        0x00650079002000520065006c0061007900511008000101020000000600511008 ,
                        0x00020102000000000051100800030102000000000006100800ffff0500050000 ,
                        0x00331000005f1002000000341000004510020000003410000003100c00030001 ,
                        0x000400040001000000331000005110080000010200000007000d101a0000000b ,
                        0x0134007800350030006d002000520065006c0061007900511008000101020000 ,
                        0x00070051100800020102000000000051100800030102000000000006100800ff ,
                        0xff060006000000331000005f1002000000341000004510020000003410000003 ,
                        0x100c00030001000400040001000000331000005110080000010200000008000d ,
                        0x10200000000e01350030006d0020004200610063006b007300740072006f006b ,
                        0x0065005110080001010200000008005110080002010200000000005110080003 ,
                        0x0102000000000006100800ffff070007000000331000005f1002000000341000 ,
                        0x004510020000003410000003100c000300010004000400010000003310000051 ,
                        0x10080000010200000009000d10240000001001350030006d0020004200720065 ,
                        0x006100730074007300740072006f006b00650051100800010102000000090051 ,
                        0x100800020102000000000051100800030102000000000006100800ffff080008 ,
                        0x000000331000005f1002000000341000004510020000003410000003100c0003 ,
                        0x000100040004000100000033100000511008000001020000000a000d101e0000 ,
                        0x000d01350030006d00200042007500740074006500720066006c007900511008 ,
                        0x000101020000000a005110080002010200000000005110080003010200000000 ,
                        0x0006100800ffff090009000000331000005f1002000000341000004510020000 ,
                        0x003410000003100c000300010004000400010000003310000051100800000102 ,
                        0x0000000b000d101e0000000d01350030006d0020004600720065006500730074 ,
                        0x0079006c006500511008000101020000000b0051100800020102000000000051 ,
                        0x100800030102000000000006100800ffff0a000a000000331000005f10020000 ,
                        0x00341000004510020000003410000003100c0003000100040004000100000033 ,
                        0x100000511008000001020000000c000d10180000000a014b00690063006b0020 ,
                        0x0042006f00610072006400511008000101020000000c00511008000201020000 ,
                        0x00000051100800030102000000000006100800ffff0b000b000000331000005f ,
                        0x1002000000341000004510020000003410000044100400190000002410020002 ,
                        0x00251020000202010000000000e2ffffffc3ffffff0000000000000000b1004d ,
                        0x0040050000331000004f10140002000200000000000000000000000000000000 ,
                        0x0026100200030051100800000102000000000034100000241002000300251020 ,
                        0x000202010000000000e2ffffffc3ffffff0000000000000000b1004d00400500 ,
                        0x00331000004f1014000200020000000000000000000000000000000000261002 ,
                        0x00030051100800000102000000000034100000461002000100411012000000a6 ,
                        0x00000087010000680b0000230c0000331000004f101400020002000000000009 ,
                        0x0100000e0c0000ee0d00001d1012000000000000000000000000000000000000 ,
                        0x0033100000201008000100010001000100621012000000000001000000010000 ,
                        0x00000000006f001e101e00020003010000000000000000000000000000000000 ,
                        0x00000023004d000000341000001d101200010000000000000000000000000000 ,
                        0x000000331000001f102a00000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000001f011e101e0002000301000000 ,
                        0x000000000000000000000000000000000023004d000000341000001410140000 ,
                        0x0000000000000000000000000000000000000033100000171006000000960000 ,
                        0x0022100a0000000000000000000f00151014002c0c0000000000003f030000a0 ,
                        0x0f000003010200331000004f101400050001002b0c0000f4ffffff69000000f5 ,
                        0x000000251020000202010000000000e2ffffffc3ffffff0000000000000000b1 ,
                        0x004d00c0160000331000004f1014000200020000000000000000000000000000 ,
                        0x0000002610020003005110080000010200000000003410000032100400000002 ,
                        0x003310000007100c00000000000500000008004d000a101000ffffff00000000 ,
                        0x00000000004e004d003410000034100000341000003410000025102000020201 ,
                        0x00000000007206000056000000bc0200000101000081404d00a0190000331000 ,
                        0x004f101400020002000000000000000000760000001500000051100800000102 ,
                        0x00000000000d102e000000150154006f00740061006c00200050006f0069006e ,
                        0x007400730020006200790020004500760065006e007400271006000100000000 ,
                        0x00341000003410000000020e00000000000400000000000c0000000a00000000 ,
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e00040000002d010000040000002d010a00040000002d0102000500 ,
                        0x00000102ffffff00050000000902000000000700000016044d019f0204001102 ,
                        0x040000002d010700040000002d010800040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010a00040000002d0102000500 ,
                        0x00000102ffffff00050000000902000000000700000016044d019f0204001102 ,
                        0x09000000fa02000001000000000000002200040000002d01090004000000f001 ,
                        0x020007000000fc020000a0e0e000605c040000002d01020004000000f0010a00 ,
                        0x050000000902ffffff00050000000102a0e0e0000400000004010d0004000000 ,
                        0x02010200070000001b0469001c02620015020500000009020000000005000000 ,
                        0x0102ffffff000400000002010100040000002e0118001c000000320a69001f02 ,
                        0x0e0000003130306d20467265657374796c650600060006000b00030006000500 ,
                        0x0700070007000400060003000700040000002e010000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016044d019f0204001102040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016044d019f020400110209000000fa0200000100 ,
                        0x0000000000002200040000002d010a0004000000f001090007000000fc020000 ,
                        0x60008000605c040000002d01090004000000f0010200050000000902ffffff00 ,
                        0x050000000102600080000400000004010d000400000002010200070000001b04 ,
                        0x84001c027d00150205000000090200000000050000000102ffffff0004000000 ,
                        0x02010100040000002e01180028000000320a84001f02160000003230306d2049 ,
                        0x6e646976696475616c204d65646c65790600060006000b000300030007000700 ,
                        0x030006000300070007000600030003000a000700070003000700060004000000 ,
                        0x2e010000040000002d010700040000002d010800040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000040000002d010900040000002d01 ,
                        0x0a00050000000102ffffff00050000000902000000000700000016044d019f02 ,
                        0x04001102040000002d010700040000002d010800040000002d01010004000000 ,
                        0x2701ffff030000001e00040000002d010000040000002d010900040000002d01 ,
                        0x0a00050000000102ffffff00050000000902000000000700000016044d019f02 ,
                        0x0400110209000000fa02000001000000000000002200040000002d0102000400 ,
                        0x0000f0010a0007000000fc020000ff808000605c040000002d010a0004000000 ,
                        0xf0010900050000000902ffffff00050000000102ff8080000400000004010d00 ,
                        0x0400000002010200070000001b049f001c029800150205000000090200000000 ,
                        0x050000000102ffffff000400000002010100040000002e01180022000000320a ,
                        0x9f001f0212000000347835306d204d65646c65792052656c6179060006000600 ,
                        0x06000b0003000a00070007000300070006000300070007000300060006000400 ,
                        0x00002e010000040000002d010700040000002d010800040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d010a0004000000 ,
                        0x2d010200050000000102ffffff00050000000902000000000700000016044d01 ,
                        0x9f0204001102040000002d010700040000002d010800040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d010a0004000000 ,
                        0x2d010200050000000102ffffff00050000000902000000000700000016044d01 ,
                        0x9f020400110209000000fa02000001000000000000002200040000002d010900 ,
                        0x04000000f001020007000000fc0200000080c000605c040000002d0102000400 ,
                        0x0000f0010a00050000000902ffffff000500000001020080c000040000000401 ,
                        0x0d000400000002010200070000001b04bb001c02b40015020500000009020000 ,
                        0x0000050000000102ffffff000400000002010100040000002e01180018000000 ,
                        0x320abb001f020b000000347835306d2052656c61790006000600060006000b00 ,
                        0x030007000700030006000600040000002e010000040000002d01070004000000 ,
                        0x2d010800040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000040000002d010200040000002d010900050000000102ffffff0005000000 ,
                        0x0902000000000700000016044d019f0204001102040000002d01070004000000 ,
                        0x2d010800040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000040000002d010200040000002d010900050000000102ffffff0005000000 ,
                        0x0902000000000700000016044d019f020400110209000000fa02000001000000 ,
                        0x000000002200040000002d010a0004000000f001090007000000fc020000c0c0 ,
                        0xff00605c040000002d01090004000000f0010200050000000902ffffff000500 ,
                        0x00000102c0c0ff000400000004010d000400000002010200070000001b04d600 ,
                        0x1c02cf00150205000000090200000000050000000102ffffff00040000000201 ,
                        0x0100040000002e0118001c000000320ad6001f020e00000035306d204261636b ,
                        0x7374726f6b65060006000b000300070006000600070007000400050007000700 ,
                        0x0700040000002e010000040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010900 ,
                        0x040000002d010a00050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f0204001102040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010900 ,
                        0x040000002d010a00050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f020400110209000000fa0200000100000000000000220004000000 ,
                        0x2d01020004000000f0010a00040000002d01030004000000f001090005000000 ,
                        0x0902ffffff00050000000102000080000400000004010d000400000002010200 ,
                        0x070000001b04f1001c02ea00150205000000090200000000050000000102ffff ,
                        0xff000400000002010100040000002e0118001f000000320af1001f0210000000 ,
                        0x35306d204272656173747374726f6b65060006000b0003000700050007000600 ,
                        0x07000400070004000500070007000700040000002e010000040000002d010700 ,
                        0x040000002d010800040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010300040000002d010200050000000102ffffff00 ,
                        0x050000000902000000000700000016044d019f0204001102040000002d010700 ,
                        0x040000002d010800040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010300040000002d010200050000000102ffffff00 ,
                        0x050000000902000000000700000016044d019f020400110209000000fa020000 ,
                        0x01000000000000002200040000002d01090004000000f0010200040000002d01 ,
                        0x0400050000000902ffffff00050000000102ff00ff000400000004010d000400 ,
                        0x000002010200070000001b040c011c0205011502050000000902000000000500 ,
                        0x00000102ffffff000400000002010100040000002e0118001b000000320a0c01 ,
                        0x1f020d00000035306d20427574746572666c7900060006000b00030007000700 ,
                        0x0400040007000500040003000600040000002e010000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016044d019f0204001102040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010400040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016044d019f020400110209000000fa0200000100 ,
                        0x0000000000002200040000002d01020004000000f0010900040000002d010500 ,
                        0x050000000902ffffff00050000000102ffff00000400000004010d0004000000 ,
                        0x02010200ffffff00050000000902000000000700000016044d019f0204001102 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010200040000002d0106000500 ,
                        0x00000102ffffff00050000000902000000000700000016044d019f0204001102 ,
                        0x09000000fa02000001000000000000002200040000002d01040004000000f001 ,
                        0x060007000000fc020000802060008020040000002d01060004000000f0010200 ,
                        0x050000000902ffffff00050000000102802060000400000004010d0004000000 ,
                        0x02010200070000001b046b003b02640034020500000009020000000005000000 ,
                        0x0102ffffff000400000002010100040000002e01180018000000320a6b003e02 ,
                        0x0b0000003230306d20537072696e74000600060006000b000300070007000500 ,
                        0x030007000400040000002e010000040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010600040000002d010400050000000102ffffff0005000000090200000000 ,
                        0x0700000016044d019f0204001102040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010600040000002d010400050000000102ffffff0005000000090200000000 ,
                        0x0700000016044d019f020400110209000000fa02000001000000000000002200 ,
                        0x040000002d01020004000000f001040007000000fc020000ffffc00080200400 ,
                        0x00002d01040004000000f0010600050000000902ffffff00050000000102ffff ,
                        0xc0000400000004010d000400000002010200070000001b04ad003b02a6003402 ,
                        0x05000000090200000000050000000102ffffff00040000000201010004000000 ,
                        0x2e0118000d000000320aad003e02040000003830306d0600060006000b000400 ,
                        0x00002e010000040000002d010300040000002d010500040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01040004000000 ,
                        0x2d010200050000000102ffffff00050000000902000000000700000016044d01 ,
                        0x9f0204001102040000002d010300040000002d010500040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01040004000000 ,
                        0x2d010200050000000102ffffff00050000000902000000000700000016044d01 ,
                        0x9f020400110209000000fa02000001000000000000002200040000002d010600 ,
                        0x04000000f001020007000000fc020000a0e0e0008020040000002d0102000400 ,
                        0x0000f0010400050000000902ffffff00050000000102a0e0e000040000000401 ,
                        0x0d000400000002010200070000001b04ee003b02e70034020500000009020000 ,
                        0x0000050000000102ffffff000400000002010100040000002e01180015000000 ,
                        0x320aee003e020900000048696768204a756d7000070003000700070003000600 ,
                        0x07000b000700040000002e010000040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010200040000002d010600050000000102ffffff0005000000090200000000 ,
                        0x0700000016044d019f0204001102040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010200040000002d010600050000000102ffffff0005000000090200000000 ,
                        0x0700000016044d019f020400110209000000fa02000001000000000000002200 ,
                        0x040000002d01040004000000f001060007000000fc0200006000800080200400 ,
                        0x00002d01060004000000f0010200050000000902ffffff000500000001026000 ,
                        0x80000400000004010d000400000002010200070000001b0430013b0229013402 ,
                        0x05000000090200000000050000000102ffffff00040000000201010004000000 ,
                        0x2e01180015000000320a30013e02090000004c6f6e67204a756d700007000700 ,
                        0x070007000300060007000b000700040000002e010000040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010400050000000102ffffff000500 ,
                        0x00000902000000000700000016044d019f0204001102040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010400050000000102ffffff000500 ,
                        0x00000902000000000700000016045101ac020000000007000000fc0200000000 ,
                        0x00000000040000002d01020004000000f0010600040000002d01030004000000 ,
                        0xf0010400040000002701ffff050000000c025101ac02030000001e0005000000 ,
                        0x0102ffffff0005000000090200000000040000002701ffff050000000b020000 ,
                        0x0000030000001e00050000000102ffffff000500000009020000000004000000 ,
                        0x2701ffff030000000000000000000000000000000000000000000000f9121aa0 ,
                        0x0c004000c031de02000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000070000001b0428011c02210115020500000009020000000005000000 ,
                        0x0102ffffff000400000002010100040000002e0118001b000000320a28011f02 ,
                        0x0d00000035306d20467265657374796c6500060006000b000300060005000700 ,
                        0x070007000400060003000700040000002e010000040000002d01070004000000 ,
                        0x2d010800040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000040000002d010500040000002d010200050000000102ffffff0005000000 ,
                        0x0902000000000700000016044d019f0204001102040000002d01070004000000 ,
                        0x2d010800040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000040000002d010500040000002d010200050000000102ffffff0005000000 ,
                        0x0902000000000700000016044d019f020400110209000000fa02000001000000 ,
                        0x000000002200040000002d01090004000000f0010200040000002d0106000500 ,
                        0x00000902ffffff0005000000010200ffff000400000004010d00040000000201 ,
                        0x0200070000001b0443011c023c01150205000000090200000000050000000102 ,
                        0xffffff000400000002010100040000002e01180016000000320a43011f020a00 ,
                        0x00004b69636b20426f6172640700030006000700030007000700060005000700 ,
                        0x040000002e010000040000002d010700040000002d010800040000002d010100 ,
                        0x040000002701ffff030000001e00040000002d010000040000002d0106000400 ,
                        0x00002d010900050000000102ffffff0005000000090200000000070000001604 ,
                        0x4d019f0204001102040000002d010700040000002d010800040000002d010100 ,
                        0x040000002701ffff030000001e00040000002d010000040000002d0106000400 ,
                        0x00002d010900050000000102ffffff0005000000090200000000070000001604 ,
                        0x5101ac020000000007000000fc020000000000000000040000002d0102000400 ,
                        0x00002d01070004000000f0010900040000002701ffff050000000c025101ac02 ,
                        0x030000001e00050000000102ffffff0005000000090200000000040000002701 ,
                        0xffff050000000b0200000000030000001e00050000000102ffffff0005000000 ,
                        0x090200000000040000002701ffff030000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0xb2460000d422000008230000010009000003841100000b002800000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c025101ac02030000001e00040000002701ffff030000001e00 ,
                        0x040000002701ffff050000000b0200000000030000001e00050000000102ffff ,
                        0xff000500000009020000000010000000fb02f5ff000000000000bc0200000000 ,
                        0x00000022417269616c000000040000002d010000040000002d01000004000000 ,
                        0x2d01000010000000fb021000070000000000bc02000000000102022253797374 ,
                        0x656d0000040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000050000000102ffffff00050000000902000000000700000016045101ac02 ,
                        0x00000000040000002d010100040000002701ffff030000001e00040000002d01 ,
                        0x0000050000000102ffffff00050000000902000000000700000016045101ac02 ,
                        0x000000000700000015044c019e0205001202040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x000000000700000016045101ac0200000000040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x000000000700000016044d01a80204000400040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x0000000007000000160425010f0224002000040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000050000000102ffffff00050000000902 ,
                        0x00000000070000001604240110022200200009000000fa020000010000000000 ,
                        0x00002200040000002d01020007000000fc0200008080ff00426c040000002d01 ,
                        0x0300050000000902ffffff000500000001028080ff000400000004010d000400 ,
                        0x0000020102000e000000240305002800de003100de0031002301280023012800 ,
                        0xde000e00000024030500a3000201ac000201ac002301a3002301a30002010e00 ,
                        0x0000240305001e01ff002701ff00270123011e0123011e01ff000e0000002403 ,
                        0x050099010101a2010101a2012301990123019901010107000000fc0200008020 ,
                        0x60000000040000002d01040004000000f0010300050000000102802060000e00 ,
                        0x0000240305003100fd003a00fd003a002301310023013100fd000e0000002403 ,
                        0x0500ac000e01b5000e01b5002301ac002301ac000e010e000000240305002701 ,
                        0x1301300113013001230127012301270113010e00000024030500a2011101ab01 ,
                        0x1101ab012301a2012301a201110107000000fc020000ffffc000000004000000 ,
                        0x2d01030004000000f0010400050000000102ffffc0000e000000240305003a00 ,
                        0xff004300ff00430023013a0023013a00ff000e00000024030500b5000a01be00 ,
                        0x0a01be002301b5002301b5000a010e0000002403050030011501390115013901 ,
                        0x230130012301300115010e00000024030500ab010601b4010601b4012301ab01 ,
                        0x2301ab01060107000000fc020000a0e0e0000000040000002d01040004000000 ,
                        0xf0010300050000000102a0e0e0000e000000240305004300df004c00df004c00 ,
                        0x2301430023014300df000e00000024030500be001a01c7001a01c7002301be00 ,
                        0x2301be001a010e000000240305003901fb004201fb0042012301390123013901 ,
                        0xfb000e00000024030500b4011801bd011801bd012301b4012301b40118010700 ,
                        0x0000fc020000600080000000040000002d01030004000000f001040005000000 ,
                        0x0102600080000e000000240305004c00e4005500e400550023014c0023014c00 ,
                        0xe4000e00000024030500c700fb00d000fb00d0002301c7002301c700fb000e00 ,
                        0x0000240305004201f0004b01f0004b012301420123014201f0000e0000002403 ,
                        0x0500bd01fe00c601fe00c6012301bd012301bd01fe0007000000fc020000c0c0 ,
                        0xff000000040000002d01040004000000f0010300050000000102c0c0ff000e00 ,
                        0x00002403050068006a0071006a00710023016800230168006a000e0000002403 ,
                        0x0500e3008600ec008600ec002301e3002301e30086000e000000240305005e01 ,
                        0xa5006701a500670123015e0123015e01a5000e00000024030500d9018000e201 ,
                        0x8000e2012301d9012301d901800007000000fc02000000008000000004000000 ,
                        0x2d01030004000000f0010400050000000102000080000e000000240305007100 ,
                        0x74007a0074007a00230171002301710074000e00000024030500ec006300f500 ,
                        0x6300f5002301ec002301ec0063000e0000002403050067017a0070017a007001 ,
                        0x23016701230167017a000e00000024030500e2016c00eb016c00eb012301e201 ,
                        0x2301e2016c0007000000fc020000ff00ff000000040000002d01040005000000 ,
                        0x0102ff00ff000e000000240305007a00e9008300e900830023017a0023017a00 ,
                        0xe9000e00000024030500f500ed00fe00ed00fe002301f5002301f500ed000e00 ,
                        0x0000240305007001be007901be0079012301700123017001be000e0000002403 ,
                        0x0500eb01e200f401e200f4012301eb012301eb01e20007000000fc020000ffff ,
                        0x00000000040000002d010500050000000102ffff00000e000000240305008300 ,
                        0x39008c0039008c00230183002301830039000e00000024030500fe0034000701 ,
                        0x340007012301fe002301fe0034000e0000002403050079014b0082014b008201 ,
                        0x23017901230179014b000e00000024030500f4015b00fd015b00fd012301f401 ,
                        0x2301f4015b0007000000fc02000000ffff000000040000002d01060005000000 ,
                        0x010200ffff000e000000240305008c00010195000101950023018c0023018c00 ,
                        0x01010e000000240305000701fd001001fd0010012301070123010701fd000e00 ,
                        0x000024030500820100018b0100018b01230182012301820100010e0000002403 ,
                        0x0500fd0105010602050106022301fd012301fd01050109000000fa0200000000 ,
                        0x0000000000002200040000002d01070007000000fc020000ffffff0000000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d01020005000000010200ffff000500 ,
                        0x00000902ffffff0007000000160425010f0224002000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d01020005000000010200ffff000500 ,
                        0x00000902ffffff000700000016045101ac020000000009000000fa0200000100 ,
                        0x0000000000002200040000002d01090004000000f00102000500000014022500 ,
                        0x2100050000000102ffffff000400000004010d00040000000201010005000000 ,
                        0x13022301210005000000140223011e0005000000130223012100050000001402 ,
                        0xf0001e00050000001302f0002100050000001402bd001e00050000001302bd00 ,
                        0x21000500000014028b001e000500000013028b00210005000000140258001e00 ,
                        0x0500000013025800210005000000140225001e00050000001302250021000500 ,
                        0x000014022301210005000000130223010d020500000014022601210005000000 ,
                        0x13022301210005000000140226019c0005000000130223019c00050000001402 ,
                        0x2601170105000000130223011701050000001402260192010500000013022301 ,
                        0x920105000000140226010d0205000000130223010d02040000002d0100000400 ,
                        0x00002d010000040000002d010700040000002d010800040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01060004000000 ,
                        0x2d010900050000000102ffffff00050000000902ffffff000700000016042200 ,
                        0x92010b001a01050000000902000000000400000004010d000400000002010100 ,
                        0x27000000320a0f001d0115000000546f74616c20506f696e7473206279204576 ,
                        0x656e740007000700040006000300030007000700030007000400070003000700 ,
                        0x0600030006000600070007000400040000002d010000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016045101ac0200000000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016045101ac0200000000040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016045201ad02ffffffff0400000004010d000400 ,
                        0x00000201010009000000320a1c01130001000000300006000a000000320ae900 ,
                        0x0d00020000003530060006000c000000320ab600070003000000313030000600 ,
                        0x060006000c000000320a8400070003000000313530000600060006000c000000 ,
                        0x320a5100070003000000323030000600060006000c000000320a1e0007000300 ,
                        0x000032353000060006000600040000002d010700040000002d01080004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0600040000002d010900050000000102ffffff00050000000902000000000700 ,
                        0x000016045101ac0200000000040000002d010700040000002d01080004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0600040000002d010900050000000102ffffff00050000000902000000000700 ,
                        0x000016045201ad02ffffffff0400000004010d0004000000020101000f000000 ,
                        0x320a2c014e00050000004153484552000800070007000600070012000000320a ,
                        0x2c01c200070000004550485241494d000600070007000700080003000a000f00 ,
                        0x0000320a2c014401050000004a5544414800060007000700080007000d000000 ,
                        0x320a2c01c401040000004c4556490700060008000300040000002d0107000400 ,
                        0x00002d010800040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010600040000002d010900050000000102ffffff000500 ,
                        0x00000902000000000700000016045101ac020000000009000000fa0205000000 ,
                        0x0000000000002200040000002d01020004000000f0010900040000002d010700 ,
                        0x040000002d010800040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010600040000002d010200050000000102ffffff00 ,
                        0x050000000902000000000700000016044d019f0204001102040000002d010700 ,
                        0x040000002d010800040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010600040000002d010200050000000102ffffff00 ,
                        0x050000000902000000000700000016044d019f020400110209000000fa020000 ,
                        0x01000000000000002200040000002d01090004000000f001020007000000fc02 ,
                        0x00008080ff00605c040000002d010200050000000902ffffff00050000000102 ,
                        0x8080ff000400000004010d000400000002010200070000001b0417001c021000 ,
                        0x150205000000090200000000050000000102ffffff0004000000020101000400 ,
                        0x00002e0118001e000000320a17001f020f0000003130306d204261636b737472 ,
                        0x6f6b65000600060006000b000300070006000600070007000400050007000700 ,
                        0x0700040000002e010000040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010200 ,
                        0x040000002d010900050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f0204001102040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010200 ,
                        0x040000002d010900050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f020400110209000000fa0200000100000000000000220004000000 ,
                        0x2d010a0004000000f001090007000000fc02000080206000605c040000002d01 ,
                        0x090004000000f0010200050000000902ffffff00050000000102802060000400 ,
                        0x000004010d000400000002010200070000001b0432001c022b00150205000000 ,
                        0x090200000000050000000102ffffff000400000002010100040000002e011800 ,
                        0x21000000320a32001f02110000003130306d204272656173747374726f6b6500 ,
                        0x0600060006000b00030007000500070006000700040007000400050007000700 ,
                        0x0700040000002e010000040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010900 ,
                        0x040000002d010a00050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f0204001102040000002d010700040000002d010800040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010900 ,
                        0x040000002d010a00050000000102ffffff000500000009020000000007000000 ,
                        0x16044d019f020400110209000000fa0200000100000000000000220004000000 ,
                        0x2d01020004000000f0010a0007000000fc020000ffffc000605c040000002d01 ,
                        0x0a0004000000f0010900050000000902ffffff00050000000102ffffc0000400 ,
                        0x000004010d000400000002010200070000001b044e001c024700150205000000 ,
                        0x090200000000050000000102ffffff000400000002010100040000002e011800 ,
                        0x1c000000320a4e001f020e0000003130306d20427574746572666c7906000600 ,
                        0x06000b000300070007000400040007000500040003000600040000002e010000 ,
                        0x040000002d010700040000002d010800040000002d010100040000002701ffff ,
                        0x030000005f2d3b5f2d405f2d1e041c00a400170000222422232c2323305f293b ,
                        0x5c28222422232c2323305c291e042100a5001c0000222422232c2323305f293b ,
                        0x5b5265645d5c28222422232c2323305c291e042200a6001d0000222422232c23 ,
                        0x23302e30305f293b5c28222422232c2323302e30305c291e042700a700220000 ,
                        0x222422232c2323302e30305f293b5b5265645d5c28222422232c2323302e3030 ,
                        0x5c291e043700a8003200005f282224222a20232c2323305f293b5f282224222a ,
                        0x205c28232c2323305c293b5f282224222a20222d225f293b5f28405f291e042e ,
                        0x00a9002900005f282a20232c2323305f293b5f282a205c28232c2323305c293b ,
                        0x5f282a20222d225f293b5f28405f291e043f00aa003a00005f282224222a2023 ,
                        0x2c2323302e30305f293b5f282224222a205c28232c2323302e30305c293b5f28 ,
                        0x2224222a20222d223f3f5f293b5f28405f291e043600ab003100005f282a2023 ,
                        0x2c2323302e30305f293b5f282a205c28232c2323302e30305c293b5f282a2022 ,
                        0x2d223f3f01000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000005f293b5f28405f2931001a00a0000100ff7fbc020000000000000501 ,
                        0x41007200690061006c0031001a00a0000100ff7fbc0200000000000005014100 ,
                        0x7200690000000000
                    End
                    RowSourceType ="Table/Query"
                    RowSource ="TRANSFORM Sum([HousePoints-Total-Event-F].SumOfPoints) AS SumOfSumOfPoints SELEC"
                        "T [HousePoints-Total-Event-F].H_Code FROM [HousePoints-Total-Event-F] GROUP BY ["
                        "HousePoints-Total-Event-F].H_Code PIVOT [HousePoints-Total-Event-F].ET_Des;"
                    Class ="MSGraph.Chart.5"
                    OLEClass ="Microsoft Graph 5.0"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
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
'Option Explicit

' Generate HTML Variables and Constants
Dim sHTML As String, rHTML As String, PageNum As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String, GenerateHTML As Boolean, aIndex As Integer
Dim ExportOleChart As Boolean

Dim HTM() As HTMarrayType

Const ReportTitle = "Overall Results - By Events"
Const repName = "ev" ' Keep to 4 letters or less (and unique from all other reports



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
        Call SpaceIndent(rHTML, 2)
        Call Text(rHTML, "", "", Me!Place)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_NAme)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_Code)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        If Not IsNull(Me!SumOfPoints) Then Call Text(rHTML, "", "", Me!SumOfPoints)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        'Debug.Print "Detail - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
        Call AddToArray(Me!ET_Des, rDetail, rHTML)
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
    rHTML = rHTML & Heading(3, "Event: " & Me!ET_Des, 3)
    Call CellEnd(rHTML)
    
    Call RowEnd(rHTML)
    
    ' *** Create general record header ***
    Call RowStart(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "PLACE")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "40%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "COMPETITOR")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "40%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TOTAL")
    Call CellEnd(rHTML)

    Call RowEnd(rHTML)

    'Debug.Print "GH - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
    Call AddToArray(Me!ET_Des, rGroupHeader, rHTML)

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    'Debug.Print "GF - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
    
    '*** HTML Generation Code Start ***
    If GenerateHTML And FormatCount = 1 Then
        
        rHTML = ""
        Call RowStart(rHTML)
    
        Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        'Debug.Print "GF - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
        Call AddToArray(Me!ET_Des, rGroupFooter, rHTML)

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
        Call AddToArray(Me!ET_Des, rPageFooter, rHTML)
        
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
        
        Call AddToArray(Me!ET_Des, rPageHeader, rHTML)

    End If

End Sub

Private Sub Report_Close()
On Error GoTo Report_Close_Err

    If GenerateHTML Then
        Dim eHTML As String, AlleHTML As String, sEvents   As String
        GenerateHTML = False
        rHTML = ""
        Call TableEnd(rHTML)
    
        Debug.Print "RF - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
        Call AddToArray(Me!ET_Des, False, rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "10%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "90%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "EVENT(S)")
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
                Call CreateHTMLfile(repName & OldPg & ".htm", Template, rHTML, PrevPage, NextPage, Heading(3, ReportTitle & "  - Page " & OldPg, 0), ReportHead)
                rHTML = ""
                
                ' *** Create summary record ***
                If OldPg Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
                
                Call RowStart(eHTML)
    
                Call CellStart(eHTML, "Center", "", "20%", BGcolor, 1)
                eHTML = eHTML & LinkStart(repName & OldPg & ".htm")
                Call Text(eHTML, "", "", Str(OldPg))
                eHTML = eHTML & LinkEnd()
                Call CellEnd(eHTML)
    
                Call CellStart(eHTML, "", "", "80%", BGcolor, 1)
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
        sHTML = sHTML & " <p align=""center""><img border=""0"" src=""events.jpg"" </p>"
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead)


    End If
    
  If ExportOleChart Then
    Dim oleGraph As Object
    HTMLFileLocation = DLookup("[HTMLlocation]", "MiscHTML")
    FileLocation = HTMLFileLocation & "\events.jpg"
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
    GenerateHTML = GlobalGenerateHTML
    ExportOleChart = GlobalGenerateHTML
    
    If GenerateHTML Then
        aIndex = 0
        PleaseWaitMsg = "Preparing HTML for """ & ReportTitle & """.  Please wait..."
        DoCmd.RunMacro "ShowPleaseWait"
    End If
    
    PageNum = 0
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
    ' ***************************


End Sub
