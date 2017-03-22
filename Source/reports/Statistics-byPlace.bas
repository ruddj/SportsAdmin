Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10649
    ItemSuffix =145
    Left =495
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xe89bd105efe5e140
    End
    RecordSource ="TRANSFORM Count([CompEvents-NewPlaces].NewPlace) AS CountOfNewPlace SELECT House"
        ".H_NAme, Sum([CompEvents-NewPlaces].Points) AS SumOfPoints FROM House INNER JOIN"
        " (EventType INNER JOIN ((Competitors INNER JOIN [CompEvents-NewPlaces] ON Compet"
        "itors.PIN = [CompEvents-NewPlaces].PIN) INNER JOIN Events ON [CompEvents-NewPlac"
        "es].E_Code = Events.E_Code) ON EventType.ET_Code = Events.ET_Code) ON House.H_Co"
        "de = Competitors.H_Code WHERE (((House.Include)=Yes) AND ((EventType.Include)=Ye"
        "s) AND ((EventType.Flag)=Yes) AND ((Events.Include)=Yes)) GROUP BY House.H_NAme,"
        " House.H_Code, House.Include ORDER BY House.H_NAme PIVOT [CompEvents-NewPlaces]."
        "NewPlace In (1,2,3,4,5,6,7,8,9,10,\"Other\");"
    Caption ="Number of Place Gained"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000992900008c01000001000000 ,
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
            ControlSource ="H_NAme"
        End
        Begin PageHeader
            Height =1700
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10596
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
                    Width =10536
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Left =120
                    Top =735
                    Width =7095
                    Height =390
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="Statistical Summary of the Number of Places Gained"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3299
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text109"
                    Caption ="1"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =9139
                    Top =1303
                    Width =1305
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Total Points"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =56
                    Top =1643
                    Width =10542
                    Name ="Line112"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3809
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text121"
                    Caption ="2"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4319
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text122"
                    Caption ="3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4839
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text123"
                    Caption ="4"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =8402
                    Top =1303
                    Width =600
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text124"
                    Caption ="Other"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5340
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text125"
                    Caption ="5"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5851
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text126"
                    Caption ="6"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6361
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text127"
                    Caption ="7"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6871
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text128"
                    Caption ="8"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7391
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text129"
                    Caption ="9"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7892
                    Top =1303
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text130"
                    Caption ="10"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =68
                    Top =1303
                    Width =645
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Text144"
                    Caption ="TEAM"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =411
            Name ="Detail1"
            Begin
                Begin Line
                    Left =56
                    Top =396
                    Width =10593
                    Name ="Line117"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3349
                    Width =456
                    Height =285
                    FontSize =10
                    Name ="Field131"
                    ControlSource ="1"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3860
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Field133"
                    ControlSource ="2"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4370
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Field134"
                    ControlSource ="3"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4881
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Field135"
                    ControlSource ="4"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5390
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Field136"
                    ControlSource ="5"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5901
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="Field137"
                    ControlSource ="6"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =6411
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="Field138"
                    ControlSource ="7"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =6922
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =7
                    Name ="Field139"
                    ControlSource ="8"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7431
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="Field140"
                    ControlSource ="9"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7942
                    Width =456
                    Height =285
                    FontSize =10
                    TabIndex =9
                    Name ="Field141"
                    ControlSource ="10"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =8452
                    Width =621
                    Height =285
                    FontSize =10
                    TabIndex =10
                    Name ="Field142"
                    ControlSource ="Other"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =9189
                    Width =1296
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="Field143"
                    ControlSource ="SumOfPoints"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =226
                    Width =2961
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="Field98"
                    ControlSource ="H_NAme"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =56
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =6059
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =5669
                    Width =1131
                    Height =390
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
                    Locked = NotDefault
                    SizeMode =3
                    RowSourceTypeInt =2
                    Left =1133
                    Top =113
                    Width =8677
                    Height =5371
                    TabIndex =1
                    Name ="Field113"
                    OleData = Begin
                        0x003a0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0500000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
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
                        0x00000046000000000000000000000000a0c4a6c496eebc0108000000800c0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000063000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002c00000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000690b0000 ,
                        0x0000000004000000fdfffffffffffffffffffffffeffffffffffffffffffffff ,
                        0xfeffffff0b000000feffffff100000000c0000000d0000000e0000000f000000 ,
                        0x0900000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a000000feffffffffffffffffffffffffffffffffffffff ,
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
                        0x0000004600000000000000000000000020f737b13067c0011b000000800c0000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000068000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002c00000014000000 ,
                        0x0000000042006f006f006b000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a0002010200000001000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000690b0000 ,
                        0x00000000ffffffffffffffff04000000fdfffffffefffffffeffffff09000000 ,
                        0xfffffffffffffffffeffffff100000000c0000000d0000000e00000006000000 ,
                        0xffffffff11000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a000000feffffff0b000000ffffffffffffffffffffffff ,
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
                        0xffffffff03004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000003000000038000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000180002010300000004000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000a0000001a170000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000001000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x28000000290000002a0000002b0000002d000000feffffff2e0000002f000000 ,
                        0x31000000fefffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff00000034100000141014000000000000000000000000000000000000 ,
                        0x000000331000001710060000009600000022100a0000000000000000000f0015 ,
                        0x101400c60e00002d000000c0000000c804000003010200331000004f10140005 ,
                        0x000100c60e00002d000000160000005100000025101a000202010000000000df ,
                        0xffffffc7ffffff0000000000000000b100331000004f10140002000200000000 ,
                        0x0000000000000000000000000051100800000102000000000034100000321004 ,
                        0x00000002003310000007100a00000000000000000009000a100c00ffffff0000 ,
                        0x00000000000000341000003410000024100200020025101a0002020100000000 ,
                        0x00dfffff01000002040000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000ffc7ffffff0000000000000000b100331000004f1014000200020000 ,
                        0x0000000000000000000000000000002610020002005110080000010200000000 ,
                        0x0034100000341000003410000025101a000202010000000000a60400004f0000 ,
                        0x004d0600001b0100008100331000004f101400020002000000000000000000f4 ,
                        0x000000190000002610020002005110080000010200000000000d102c00000029 ,
                        0x4e756d626572206f662022506c6163657322204761696e656420647572696e67 ,
                        0x2043617201000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x28000000290000002a0000002b0000002d000000feffffff2e0000002f000000 ,
                        0x31000000fefffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff0100feff030a0000ffffffff0108020000000000c000000000000046 ,
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
                        0x23302e30305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a20222d223f3f ,
                        0x5f2d3b5f2d405f2d1e041a00a40017222422232c2323305f293b5c2822242223 ,
                        0x2c23233038000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x000000006e6976616c27100600010000000000341000003410000000000a0000 ,
                        0x0008000000050000000a00000000000000000000000000000000000000000000 ,
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
                        0x00000000ffffffff030000000400000001000000ffffffff0000000000000000 ,
                        0x903f000084250000f2160000010009000003790b000007004500000000000500 ,
                        0x0000090200000000050000000102ffffff000400000004010d00040000000201 ,
                        0x0200050000000c026b016702030000001e00040000002701ffff050000000b02 ,
                        0x00000000030000001e00050000000102ffffff00050000000902000000001000 ,
                        0x0000fb02f5ff000000000000bc020000000000000022417269616c00a9810400 ,
                        0x00002d01000010000000fb021000070000000000bc0200000000010202225379 ,
                        0x7374656d006e040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016046b01 ,
                        0x670200000000040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000050000000102ffffff00050000000902000000000700000016046b01 ,
                        0x67020000000009000000fa02000001000000000000002200040000002d010200 ,
                        0x09000000fa02000000000000000000002200040000002d010300040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010200 ,
                        0x050000000102ffffff0005000000090200000000070000001604460134023a00 ,
                        0x380007000000150475005e020900410205000000140221013900040000000401 ,
                        0x0d0004005c291e041f00a5001c222422232c2323305f293b5b5265645d5c2822 ,
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
                        0xff7fbc0200000000007b05417269616c31001400c8000100ff7fbc0200000000 ,
                        0x007b05417269616c31001400a0000100ff7fbc0200000002007b05417269616c ,
                        0x31001400c8000100ff7fbc0200000000007b05417269616c3d001200dd04c003 ,
                        0x8124da16000000003e200000000085000700590300000002000a000000090808 ,
                        0x0080050080e209c907ac02020038005210040001021000331000008c00040001 ,
                        0x003d002610020003005310040000000500541004000000060055100600000000 ,
                        0x00000104000e000000000000000006485f4e416d6503000f0000000100000000 ,
                        0x000000000000f03f03000f0000000200000000000000000000004003000f0000 ,
                        0x000300000000000000000000084003000f000000040000000000000000000010 ,
                        0x4003000f0000000500000000000000000000144004000d000100000000000005 ,
                        0x415348455203000f0001000100000000000000000000394003000f0001000200 ,
                        0x0000000000000000003c4003000f000100030000000000000000000044400300 ,
                        0x0f0001000400000000000000000000394003000f000100050000000000000000 ,
                        0x0000404004000f0002000000000000074550485241494d03000f000200010000 ,
                        0x0000000000000080444003000f00020002000000000000000000003c4003000f ,
                        0x0002000300000000000000000000364003000f00020004000000000000000000 ,
                        0x00364003000f00020005000000000000000000002c4004000d00030000000000 ,
                        0x00054a5544414803000f0003000100000000000000000080404003000f000300 ,
                        0x0200000000000000000080444003000f00030003000000000000000000003640 ,
                        0x03000f0003000400000000000000000000374003000f00030005000000000000 ,
                        0x00000000354004000c0004000000000000044c45564903000f00040001000000 ,
                        0x00000000000080444003000f0004000200000000000000000000424003000f00 ,
                        0x04000300000000000000000080404003000f0004000400000000000000000080 ,
                        0x444003000f00040005000000000000000000003b40571001000159100800b400 ,
                        0x2c01eb3294203d000a0000000000f627000f00003e000e000101010001010001 ,
                        0x0001000000005810020000001d00110003010001000000010001000100010001 ,
                        0x00341000000110020000000210100000000000000000000040cd010040100133 ,
                        0x100000a000040001000100031008000300010004000400331000005110080000 ,
                        0x010200000001000d100400000001315110080001010200000001005110080002 ,
                        0x0102000000000006100800ffff0000000000003310000007100a000000000000 ,
                        0x00000001000a100c008080ff00ffffff00010001000b100200000009100c0000 ,
                        0x0080000000800002000100341000004510020000003410000003100800030001 ,
                        0x0004000400331000005110080000010200000002000d10040000000132511008 ,
                        0x00010102000000020051100800020102000000000006100800ffff0100010000 ,
                        0x0045100200000034100000031008000300010004000400331000005110080000 ,
                        0x010200000003000d100400000001335110080001010200000003005110080002 ,
                        0x0102000000000006100800ffff02000200000045100200000034100000031008 ,
                        0x000300010004000400331000005110080000010200000004000d100400000001 ,
                        0x3451100800010102000000040051100800020102000000000006100800ffff03 ,
                        0x0003000000451002000000341000000310080003000100040004003310000051 ,
                        0x10080000010200000005000d1004000000013551100800010102000000050051 ,
                        0x100800020102000000000006100800ffff040004000000451002000000341000 ,
                        0x00441003000900004610020001004110120000004401000091020000250d0000 ,
                        0x6c0b0000331000004f10140002000200b200000025020000ba0d00000a0d0000 ,
                        0x1d101200000000000000000000000000000000000000331000001e101a000200 ,
                        0x030100000000000000000000000000000000000000002300341000001d101200 ,
                        0x010000000000000000000000000000000000331000001f102a00000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00001f011e101a00020003010000000000000000000000000000000000000000 ,
                        0x230021100200010007100a00000000000000000009003410000025101a000202 ,
                        0x0100000000000d000000ee050000a5000000bc0400008902331000004f101400 ,
                        0x020002000000000000000000190000006b000000261002000200511008000001 ,
                        0x0200000000000d1013000000104e756d626572206f6620506c61636573271006 ,
                        0x0002000000000034100000141014000000000000000000000000000000000000 ,
                        0x000000331000001710060000009600000022100a0000000000000000000f0015 ,
                        0x101400c60e00002d000000c0000000c804000003010200331000004f10140005 ,
                        0x000100c60e00002d000000160000005100000025101a000202010000000000df ,
                        0xffffffc7ffffff0000000000000000b100331000004f10140002000200000000 ,
                        0x0000000000000000000000000051100800000102000000000034100000321004 ,
                        0x00000002003310000007100a00000000000000000009000a100c00ffffff0000 ,
                        0x00000000000000341000003410000024100200020025101a0002020100000000 ,
                        0x00dfffff01000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000ffc7ffffff0000000000000000b100331000004f1014000200020000 ,
                        0x0000000000000000000000000000002610020002005110080000010200000000 ,
                        0x0034100000341000003410000025101a000202010000000000a60400004f0000 ,
                        0x004d0600001b0100008100331000004f101400020002000000000000000000f4 ,
                        0x000000190000002610020002005110080000010200000000000d102c00000029 ,
                        0x4e756d626572206f662022506c6163657322204761696e656420647572696e67 ,
                        0x2043617200000201010005000000130221013302050000001402010139000500 ,
                        0x0000130201013302050000001402e0003900050000001302e000330205000000 ,
                        0x1402c0003900050000001302c0003302050000001402a0003900050000001302 ,
                        0xa000330205000000140280003900050000001302800033020500000014025f00 ,
                        0x39000500000013025f0033020500000014023f0039000500000013023f003302 ,
                        0x040000002d010300040000002d010100040000002701ffff030000001e000400 ,
                        0x00002d010000040000002d010200050000000102ffffff000500000009020000 ,
                        0x00000700000016046b0167020000000007000000150475005e02090041020400 ,
                        0x00002d010300040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200050000000102ffffff0005000000090200000000 ,
                        0x0700000016046b01670200000000040000002d010300040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01020005000000 ,
                        0x0102ffffff000500000009020000000007000000160467016302040004000400 ,
                        0x00002d010300040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200050000000102ffffff0005000000090200000000 ,
                        0x070000001604430135023e003800040000002d010300040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01020005000000 ,
                        0x0102ffffff0005000000090200000000070000001604420136023c0038000900 ,
                        0x0000fa02000001000000000000002200040000002d01040004000000f0010200 ,
                        0x07000000fc0200008080ff00ffff040000002d010200050000000902ffffff00 ,
                        0x0500000001028080ff000400000004010d0004000000020102000e0000002403 ,
                        0x05004700c0005b00c0005b004101470041014700c0000e00000024030500c600 ,
                        0x3f00d9003f00d9004101c6004101c6003f000e000000240305004401c0005801 ,
                        0xc00058014101440141014401c00007000000fc02000080206000ffff04000000 ,
                        0x2d01050004000000f0010200050000000102802060000e000000240305005b00 ,
                        0x3f006e003f006e0041015b0041015b003f000e00000024030500d9000101ed00 ,
                        0x0101ed004101d9004101d90001010e00000024030500d6018000ea018000ea01 ,
                        0x4101d6014101d601800007000000fc020000ffffc000ffff040000002d010200 ,
                        0x04000000f0010500050000000102ffffc0000e000000240305006e003f008200 ,
                        0x3f00820041016e0041016e003f000e00000024030500ed000101000101010001 ,
                        0x4101ed004101ed0001010e000000240305006b0180007f0180007f0141016b01 ,
                        0x41016b0180000e00000024030500ea018000fd018000fd014101ea014101ea01 ,
                        0x800007000000fc020000a0e0e000ffff040000002d01050004000000f0010200 ,
                        0x050000000102a0e0e0000e000000240305007f01c0009201c000920141017f01 ,
                        0x41017f01c0000e00000024030500fd01c0001102c00011024101fd014101fd01 ,
                        0xc00007000000fc02000060008000ffff040000002d01020004000000f0010500 ,
                        0x050000000102600080000e000000240305001401010127010101270141011401 ,
                        0x4101140101010e000000240305009201c000a601c000a6014101920141019201 ,
                        0xc0000e0000002403050011023f0024023f00240241011102410111023f000400 ,
                        0x00002d01030007000000fc020000ffffff000000040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d01040005000000010260008000050000000902ffffff000700 ,
                        0x00001604430135023e003800040000002d010300040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d01040005000000010260008000050000000902ffffff000700 ,
                        0x000016046b0167020000000009000000fa020000010000000000000022000400 ,
                        0x00002d01060004000000f00104000500000014023f003900050000000102ffff ,
                        0xff000400000004010d0004000000020101000500000013024101390005000000 ,
                        0x1402410136000500000013024101390005000000140221013600050000001302 ,
                        0x210139000500000014020101360005000000130201013900050000001402e000 ,
                        0x3600050000001302e0003900050000001402c0003600050000001302c0003900 ,
                        0x050000001402a0003600050000001302a0003900050000001402800036000500 ,
                        0x00001302800039000500000014025f0036000500000013025f00390005000000 ,
                        0x14023f0036000500000013023f00390005000000140241013900050000001302 ,
                        0x4101330205000000140244013900050000001302410139000500000014024401 ,
                        0xb8000500000013024101b8000500000014024401360105000000130241013601 ,
                        0x0500000014024401b5010500000013024101b501050000001402440133020500 ,
                        0x0000130241013302040000002d010300040000002d010500040000002d010100 ,
                        0x040000002701ffff030000001e00040000002d010000040000002d0102000400 ,
                        0x00002d010600050000000102ffffff00050000000902ffffff00070000001604 ,
                        0x2600ae010b00b800050000000902000000000400000004010d00040000000201 ,
                        0x010045000000320a1100be00290000004e756d626572206f662022506c616365 ,
                        0x7322204761696e656420647572696e67204361726e6976616c00070007000b00 ,
                        0x0700070005000300070004000300050007000300060006000700070005000300 ,
                        0x0800060003000700070007000300070007000500030007000700030008000600 ,
                        0x050007000300060006000300040000002d010300040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d010600050000000102ffffff00050000000902000000000700 ,
                        0x000016046b01670200000000040000002d010300040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d010600050000000102ffffff00050000000902000000000700 ,
                        0x000016046b01670200000000040000002d010300040000002d01050004000000 ,
                        0x2d010100040000002701ffff030000001e00040000002d010000040000002d01 ,
                        0x0200040000002d010600050000000102ffffff00050000000902000000000700 ,
                        0x000016046c016802ffffffff0400000004010d00040000000201010009000000 ,
                        0x320a3a012b0001000000300006000c000000320a1a01220003000000302e3500 ,
                        0x06000300060009000000320afa002b0001000000310006000c000000320ad900 ,
                        0x220003000000312e350006000300060009000000320ab9002b00010000003200 ,
                        0x06000c000000320a9900220003000000322e350006000300060009000000320a ,
                        0x79002b0001000000330006000c000000320a5800220003000000332e35000600 ,
                        0x0300060009000000320a38002b000100000034000600040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016046b01670200000000040000002d0103000400 ,
                        0x00002d010500040000002d010100040000002701ffff030000001e0004000000 ,
                        0x2d010000040000002d010200040000002d010600050000000102ffffff000500 ,
                        0x00000902000000000700000016046c016802ffffffff0400000004010d000400 ,
                        0x0000020101000f000000320a4a01670005000000415348455200080007000700 ,
                        0x0600070012000000320a4a01df00070000004550485241494d00060007000700 ,
                        0x0700080003000a000f000000320a4a016401050000004a554441480006000700 ,
                        0x0700080007000d000000320a4a01e801040000004c4556490700060008000300 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010200040000002d0106000500 ,
                        0x00000102ffffff00050000000902000000000700000016046b01670200000000 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010200040000002d0106000500 ,
                        0x00000102ffffff0005000000090200000000070000001604f70021008a000600 ,
                        0x0400000004010d00040000000201010010000000fb02f5ff000084038403bc02 ,
                        0x0000000000100022417269616c00aa81040000002d0104000f000000320af100 ,
                        0x0c00100000004e756d626572206f6620506c61636573040000002d0100000400 ,
                        0x0000f0010400040000002d010300040000002d010500040000002d0101000400 ,
                        0x00002701ffff030000001e00040000002d010000040000002d01020004000000 ,
                        0x2d010600050000000102ffffff00050000000902000000000700000016046b01 ,
                        0x67020000000007000000fc020100000000000000040000002d01040004000000 ,
                        0xf00102000400000004010d000400000002010200070000001b0476005e020900 ,
                        0x4102040000002d010300040000002d010500040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000040000002d010400040000002d010600 ,
                        0x050000000102ffffff000500000009020000000007000000160475005e020900 ,
                        0x4102040000002d010300040000002d010500040000002d010100040000002701 ,
                        0xffff030000001e00040000002d010000040000002d010400040000002d010600 ,
                        0x050000000102ffffff000500000009020000000007000000160475005e020900 ,
                        0x410209000000fa02000001000000000000002200040000002d01020004000000 ,
                        0xf001060007000000fc0200008080ff00ffff040000002d01060004000000f001 ,
                        0x0400050000000902ffffff000500000001028080ff000400000004010d000400 ,
                        0x000002010200070000001b041800500211004902050000000902000000000500 ,
                        0x00000102ffffff000400000002010100040000002e01180009000000320a1800 ,
                        0x53020100000031000600040000002e010000040000002d010300040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010600040000002d010200050000000102ffffff00050000000902 ,
                        0x0000000007000000160475005e0209004102040000002d010300040000002d01 ,
                        0x0500040000002d010100040000002701ffff030000001e00040000002d010000 ,
                        0x040000002d010600040000002d010200050000000102ffffff00050000000902 ,
                        0x0000000007000000160475005e020900410207000000fc02000080206000ffff ,
                        0x040000002d01040004000000f0010600050000000902ffffff00050000000102 ,
                        0x802060000400000004010d000400000002010200070000001b042d0050022600 ,
                        0x490205000000090200000000050000000102ffffff0004000000020101000400 ,
                        0x00002e01180009000000320a2d0053020100000032000600040000002e010000 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010400040000002d0102000500 ,
                        0x00000102ffffff000500000009020000000007000000160475005e0209004102 ,
                        0x040000002d010300040000002d010500040000002d010100040000002701ffff ,
                        0x030000001e00040000002d010000040000002d010400040000002d0102000500 ,
                        0x00000102ffffff000500000009020000000007000000160475005e0209004102 ,
                        0x07000000fc020000ffffc000ffff040000002d01060004000000f00104000500 ,
                        0x00000902ffffff00050000000102ffffc0000400000004010d00040000000201 ,
                        0x0200070000001b04420050023b00490205000000090200000000050000000102 ,
                        0xffffff000400000002010100040000002e01180009000000320a420053020100 ,
                        0x000033000600040000002e010000040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010600040000002d010200050000000102ffffff0005000000090200000000 ,
                        0x07000000160475005e0209004102040000002d010300040000002d0105000400 ,
                        0x00002d010100040000002701ffff030000001e00040000002d01000004000000 ,
                        0x2d010600040000002d010200050000000102ffffff0005000000090200000000 ,
                        0x07000000160475005e020900410207000000fc020000a0e0e000ffff04000000 ,
                        0x2d01040004000000f0010600050000000902ffffff00050000000102a0e0e000 ,
                        0x0400000004010d000400000002010200070000001b0458005002510049020500 ,
                        0x0000090200000000050000000102ffffff000400000002010100040000002e01 ,
                        0x180009000000320a580053020100000034000600040000002e01000004000000 ,
                        0x2d010300040000002d010500040000002d010100040000002701ffff03000000 ,
                        0x1e00040000002d010000040000002d010400040000002d010200050000000102 ,
                        0xffffff000500000009020000000007000000160475005e020900410204000000 ,
                        0x2d010300040000002d010500040000002d010100040000002701ffff03000000 ,
                        0x1e00040000002d010000040000002d010400040000002d010200050000000102 ,
                        0xffffff000500000009020000000007000000160475005e020900410207000000 ,
                        0xfc02000060008000ffff040000002d01060004000000f0010400050000000902 ,
                        0xffffff00050000000102600080000400000004010d0004000000020102000700 ,
                        0x00001b046d0050026600490205000000090200000000050000000102ffffff00 ,
                        0x0400000002010100040000002e01180009000000320a6d005302010000003500 ,
                        0x0600040000002e010000040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010600 ,
                        0x040000002d010200050000000102ffffff000500000009020000000007000000 ,
                        0x160475005e0209004102040000002d010300040000002d010500040000002d01 ,
                        0x0100040000002701ffff030000001e00040000002d010000040000002d010600 ,
                        0x040000002d010200050000000102ffffff000500000009020000000007000000 ,
                        0x16046b0167020000000007000000fc020000000000000000040000002d010400 ,
                        0x04000000f0010600040000002d01030004000000f0010200040000002701ffff ,
                        0x050000000c026b016702030000001e00050000000102ffffff00050000000902 ,
                        0x00000000040000002701ffff050000000b0200000000030000001e0005000000 ,
                        0x0102ffffff0005000000090200000000040000002701ffff0300000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000001f011e101a000200030100000000000000000000000000000000000000 ,
                        0x0023003410000014101400000000000000000000000000000000000000010033 ,
                        0x1000001710060000009600000022100a0000000000000000000f0016100c0005 ,
                        0x000100020003000400050024100200020025101a000202010000000000deffff ,
                        0xffc7ffffff0000000000000000b100331000004f101400020002000000000000 ,
                        0x0000000000000000000000261002000200511008000001020000000000341000 ,
                        0x003410000100feff030a0000ffffffff0308020000000000c000000000000046 ,
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
                        0x23302e30305f2d3b5c2d2a20232c2323302e30305f2d3b5f2d2a20222d223f3f ,
                        0x5f2d3b5f2d405f2d1e041a00a40017222422232c2323305f293b5c2822242223 ,
                        0x2c23233000000000
                    End
                    RowSourceType ="Table/Query"
                    RowSource ="TRANSFORM Count(CompEvents.Place) AS CountOfPlace SELECT House.H_NAme FROM Event"
                        "Type INNER JOIN ((House INNER JOIN (Competitors INNER JOIN CompEvents ON Competi"
                        "tors.PIN = CompEvents.PIN) ON House.H_Code = Competitors.H_Code) INNER JOIN Even"
                        "ts ON CompEvents.E_Code = Events.E_Code) ON EventType.ET_Code = Events.ET_Code W"
                        "HERE (((House.Include)=Yes) AND ((EventType.Include)=Yes) AND ((EventType.Flag)="
                        "Yes) AND ((Events.Include)=Yes) AND ((CompEvents.Place)>0 And (CompEvents.Place)"
                        "<6)) GROUP BY House.H_NAme, House.H_Code ORDER BY House.H_NAme, CompEvents.Place"
                        " PIVOT CompEvents.Place;"
                    Class ="MSGraph.Chart.5"
                    OLEClass ="Microsoft Graph 5.0"

                End
                Begin Line
                    Left =1814
                    Width =7818
                    Name ="Line120"
                End
            End
        End
    End
End
