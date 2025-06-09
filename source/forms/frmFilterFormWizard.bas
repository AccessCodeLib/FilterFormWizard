Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =10771
    DatasheetFontHeight =10
    ItemSuffix =136
    Left =2835
    Top =1470
    Right =13605
    Bottom =7823
    RecSrcDt = Begin
        0x668d2cd46a58e440
    End
    Caption ="ACLib FilterForm Wizard"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =6362
            BackColor =-2147483633
            Name ="Detailbereich"
            Begin
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =2345
                    Top =5394
                    Width =295
                    Height =295
                    FontSize =16
                    ForeColor =5026082
                    Name ="labCopyModulFilterStringBuilder"
                    FontName ="Marlett"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1
                    LayoutCachedLeft =2345
                    LayoutCachedTop =5394
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5689
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =2345
                    Top =4710
                    Width =295
                    Height =295
                    FontSize =16
                    ForeColor =5026082
                    Name ="labCopyModulFilterControlManager"
                    FontName ="Marlett"
                    VerticalAnchor =1
                    LayoutCachedLeft =2345
                    LayoutCachedTop =4710
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5005
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =2345
                    Top =5769
                    Width =295
                    Height =295
                    FontSize =16
                    ForeColor =5026082
                    Name ="labCopyModulSqlTools"
                    FontName ="Marlett"
                    VerticalAnchor =1
                    LayoutCachedLeft =2345
                    LayoutCachedTop =5769
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =6064
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =2345
                    Top =4371
                    Width =295
                    Height =295
                    FontSize =16
                    ForeColor =5026082
                    Name ="labCopyModules"
                    FontName ="Marlett"
                    VerticalAnchor =1
                    LayoutCachedLeft =2345
                    LayoutCachedTop =4371
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =4666
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =2748
                    Top =4210
                    Width =0
                    Height =2098
                    Name ="Linie49"
                    VerticalAnchor =1
                    LayoutCachedLeft =2748
                    LayoutCachedTop =4210
                    LayoutCachedWidth =2748
                    LayoutCachedHeight =6308
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    Name ="sysFirst"

                End
                Begin CommandButton
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7426
                    Top =5832
                    Width =3072
                    Height =375
                    FontSize =10
                    TabIndex =18
                    Name ="cmdAddFilterCodeToForm"
                    Caption ="Insert Form Code"
                    OnClick ="[Event Procedure]"
                    FontName ="Verdana"
                    Tag ="LANG:"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =7426
                    LayoutCachedTop =5832
                    LayoutCachedWidth =10498
                    LayoutCachedHeight =6207
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =315
                    Top =4740
                    Width =1991
                    Height =255
                    FontSize =9
                    Name ="labCaptionCopyModulFilterControlManager"
                    Caption ="FilterControlManager"
                    VerticalAnchor =1
                    LayoutCachedLeft =315
                    LayoutCachedTop =4740
                    LayoutCachedWidth =2306
                    LayoutCachedHeight =4995
                End
                Begin Label
                    OverlapFlags =85
                    Left =315
                    Top =5419
                    Width =1991
                    Height =255
                    FontSize =9
                    Name ="labCaptionCopyModulFilterStringBuilder"
                    Caption ="FilterStringBuilder"
                    VerticalAnchor =1
                    LayoutCachedLeft =315
                    LayoutCachedTop =5419
                    LayoutCachedWidth =2306
                    LayoutCachedHeight =5674
                End
                Begin Label
                    OverlapFlags =85
                    Left =315
                    Top =5794
                    Width =1991
                    Height =255
                    FontSize =9
                    Name ="labCaptionCopyModulSqlTools"
                    Caption ="SqlTools"
                    VerticalAnchor =1
                    LayoutCachedLeft =315
                    LayoutCachedTop =5794
                    LayoutCachedWidth =2306
                    LayoutCachedHeight =6049
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =215
                    Left =2345
                    Top =5394
                    Width =295
                    Height =295
                    TabIndex =15
                    Name ="cmdCopyModulFilterStringBuilder"
                    Caption ="Befehl69"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =2345
                    LayoutCachedTop =5394
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5689
                    Overlaps =1
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =215
                    Left =2345
                    Top =4710
                    Width =295
                    Height =295
                    TabIndex =11
                    Name ="cmdCopyModulFilterControlManager"
                    Caption ="Befehl69"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =2345
                    LayoutCachedTop =4710
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5005
                    Overlaps =1
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =215
                    Left =2345
                    Top =5769
                    Width =295
                    Height =295
                    TabIndex =17
                    Name ="cmdCopyModulSqlTools"
                    Caption ="Befehl69"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =2345
                    LayoutCachedTop =5769
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =6064
                    Overlaps =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =84
                    Top =4393
                    Width =2145
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =16737792
                    Name ="labCopyCaption"
                    Caption ="Install classes:"
                    OnMouseDown ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="LANG:"
                    VerticalAnchor =1
                    LayoutCachedLeft =84
                    LayoutCachedTop =4393
                    LayoutCachedWidth =2229
                    LayoutCachedHeight =4648
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =223
                    Left =2345
                    Top =4371
                    Width =295
                    Height =295
                    TabIndex =9
                    Name ="cmdCopyModules"
                    Caption ="Befehl69"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =2345
                    LayoutCachedTop =4371
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =4666
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2003
                    Top =165
                    Width =3402
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="cbxFormName"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Verdana"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =2003
                    LayoutCachedTop =165
                    LayoutCachedWidth =5405
                    LayoutCachedHeight =450
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =77
                            Top =165
                            Width =1920
                            Height =285
                            FontSize =10
                            Name ="labSelectForm"
                            Caption ="Filter form:"
                            FontName ="Verdana"
                            Tag ="LANG:"
                            LayoutCachedLeft =77
                            LayoutCachedTop =165
                            LayoutCachedWidth =1997
                            LayoutCachedHeight =450
                        End
                    End
                End
                Begin Subform
                    Enabled = NotDefault
                    OverlapFlags =87
                    SpecialEffect =3
                    Left =84
                    Top =895
                    Width =10605
                    Height =3255
                    TabIndex =7
                    BorderColor =-2147483630
                    Name ="sfrFilterControls"
                    SourceObject ="Form.frmFilterFormWizard_SF_Controls"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =84
                    LayoutCachedTop =895
                    LayoutCachedWidth =10689
                    LayoutCachedHeight =4150
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =84
                            Top =610
                            Width =1920
                            Height =285
                            FontSize =10
                            Name ="labFilterControls"
                            Caption ="Filter controls:"
                            Tag ="LANG:"
                            LayoutCachedLeft =84
                            LayoutCachedTop =610
                            LayoutCachedWidth =2004
                            LayoutCachedHeight =895
                        End
                    End
                End
                Begin OptionGroup
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =3120
                    Top =4705
                    Width =3742
                    Height =1304
                    TabIndex =10
                    Name ="ApplyFilterMethodOptions"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =3120
                    LayoutCachedTop =4705
                    LayoutCachedWidth =6862
                    LayoutCachedHeight =6009
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =3240
                            Top =4590
                            Width =1536
                            Height =240
                            BackColor =-2147483633
                            Name ="labUseFilterMethodOptions"
                            Caption ="ApplyFilter method"
                            Tag ="LANG:"
                            HorizontalAnchor =1
                            VerticalAnchor =1
                            LayoutCachedLeft =3240
                            LayoutCachedTop =4590
                            LayoutCachedWidth =4776
                            LayoutCachedHeight =4830
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3270
                            Top =4977
                            OptionValue =0
                            Name ="ApplyFilterMethodOptions_Opt0"
                            HorizontalAnchor =1
                            VerticalAnchor =1

                            LayoutCachedLeft =3270
                            LayoutCachedTop =4977
                            LayoutCachedWidth =3530
                            LayoutCachedHeight =5217
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3495
                                    Top =4945
                                    Width =3075
                                    Height =240
                                    Name ="labUseFilterMethodOptions0"
                                    Caption ="insert sample code"
                                    Tag ="LANG:"
                                    HorizontalAnchor =1
                                    VerticalAnchor =1
                                    LayoutCachedLeft =3495
                                    LayoutCachedTop =4945
                                    LayoutCachedWidth =6570
                                    LayoutCachedHeight =5185
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3270
                            Top =5287
                            TabIndex =1
                            OptionValue =1
                            Name ="ApplyFilterMethodOptions_Opt1"
                            HorizontalAnchor =1
                            VerticalAnchor =1

                            LayoutCachedLeft =3270
                            LayoutCachedTop =5287
                            LayoutCachedWidth =3530
                            LayoutCachedHeight =5527
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3499
                                    Top =5260
                                    Width =3049
                                    Height =240
                                    Name ="labUseFilterMethodOptions1"
                                    Caption ="filter current form"
                                    Tag ="LANG:"
                                    HorizontalAnchor =1
                                    VerticalAnchor =1
                                    LayoutCachedLeft =3499
                                    LayoutCachedTop =5260
                                    LayoutCachedWidth =6548
                                    LayoutCachedHeight =5500
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3270
                            Top =5630
                            TabIndex =2
                            OptionValue =2
                            Name ="ApplyFilterMethodOptions_Opt2"
                            HorizontalAnchor =1
                            VerticalAnchor =1

                            LayoutCachedLeft =3270
                            LayoutCachedTop =5630
                            LayoutCachedWidth =3530
                            LayoutCachedHeight =5870
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3499
                                    Top =5600
                                    Width =1114
                                    Height =240
                                    Name ="labUseFilterMethodOptions3"
                                    Caption ="Subform"
                                    Tag ="LANG:"
                                    HorizontalAnchor =1
                                    VerticalAnchor =1
                                    LayoutCachedLeft =3499
                                    LayoutCachedTop =5600
                                    LayoutCachedWidth =4613
                                    LayoutCachedHeight =5840
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListWidth =3402
                    Left =4613
                    Top =5600
                    Width =2061
                    TabIndex =16
                    Name ="cbxApplyFilterSubForm"
                    RowSourceType ="Value List"
                    ColumnWidths ="3402"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =4613
                    LayoutCachedTop =5600
                    LayoutCachedWidth =6674
                    LayoutCachedHeight =5840
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =1176
                    Top =4183
                    Width =1548
                    Height =210
                    FontSize =7
                    Name ="labCodeInstalled"
                    Caption ="already installed"
                    FontName ="Small Fonts"
                    Tag ="LANG:"
                    VerticalAnchor =1
                    LayoutCachedLeft =1176
                    LayoutCachedTop =4183
                    LayoutCachedWidth =2724
                    LayoutCachedHeight =4393
                End
                Begin OptionGroup
                    Enabled = NotDefault
                    OverlapFlags =223
                    Left =7260
                    Top =4720
                    Width =3458
                    Height =1574
                    TabIndex =12
                    Name ="FormCodeOptions"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =7260
                    LayoutCachedTop =4720
                    LayoutCachedWidth =10718
                    LayoutCachedHeight =6294
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =7376
                            Top =4590
                            Width =1395
                            Height =240
                            BackColor =-2147483633
                            Name ="labFormCodeOptions"
                            Caption ="Filter code"
                            Tag ="LANG:"
                            HorizontalAnchor =1
                            VerticalAnchor =1
                            LayoutCachedLeft =7376
                            LayoutCachedTop =4590
                            LayoutCachedWidth =8771
                            LayoutCachedHeight =4830
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =7365
                            Top =4937
                            OptionValue =1
                            Name ="FormCodeOptions_Opt1"
                            HorizontalAnchor =1
                            VerticalAnchor =1

                            LayoutCachedLeft =7365
                            LayoutCachedTop =4937
                            LayoutCachedWidth =7625
                            LayoutCachedHeight =5177
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =7594
                                    Top =4905
                                    Width =2948
                                    Height =240
                                    Name ="labFormCodeOptions1"
                                    Caption ="FilterControlManager methods"
                                    Tag ="LANG:"
                                    HorizontalAnchor =1
                                    VerticalAnchor =1
                                    LayoutCachedLeft =7594
                                    LayoutCachedTop =4905
                                    LayoutCachedWidth =10542
                                    LayoutCachedHeight =5145
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =7365
                            Top =5517
                            TabIndex =1
                            OptionValue =2
                            Name ="FormCodeOptions_Opt2"
                            HorizontalAnchor =1
                            VerticalAnchor =1

                            LayoutCachedLeft =7365
                            LayoutCachedTop =5517
                            LayoutCachedWidth =7625
                            LayoutCachedHeight =5757
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =7594
                                    Top =5490
                                    Width =2948
                                    Height =240
                                    Name ="Label130"
                                    Caption ="FilterStringBuilder methods"
                                    Tag ="LANG:"
                                    HorizontalAnchor =1
                                    VerticalAnchor =1
                                    LayoutCachedLeft =7594
                                    LayoutCachedTop =5490
                                    LayoutCachedWidth =10542
                                    LayoutCachedHeight =5730
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =8831
                    Top =4195
                    Width =1836
                    TabIndex =8
                    Name ="cbxSqlLang"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tabSqlLangFormat.SqlLang, tabSqlLangFormat.SqlDateFormat, tabSqlLangForma"
                        "t.SqlBooleanTrueString, tabSqlLangFormat.SqlWildCardString FROM tabSqlLangFormat"
                        " ORDER BY tabSqlLangFormat.SqlLang;"
                    ColumnWidths ="1701;0;0;0"
                    StatusBarText ="Select SQL dialect"
                    DefaultValue ="\"Jet/DAO\""
                    Tag ="LANG:"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =8831
                    LayoutCachedTop =4195
                    LayoutCachedWidth =10667
                    LayoutCachedHeight =4435
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =7766
                            Top =4195
                            Width =1065
                            Height =240
                            RightMargin =29
                            Name ="labSqlLang"
                            Caption ="SQL dialect:"
                            Tag ="LANG:"
                            HorizontalAnchor =1
                            VerticalAnchor =1
                            LayoutCachedLeft =7766
                            LayoutCachedTop =4195
                            LayoutCachedWidth =8831
                            LayoutCachedHeight =4435
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =2268
                    Left =8423
                    Top =308
                    Width =1659
                    TabIndex =4
                    Name ="cbxRemoveFilterCtl"
                    RowSourceType ="Value List"

                    LayoutCachedLeft =8423
                    LayoutCachedTop =308
                    LayoutCachedWidth =10082
                    LayoutCachedHeight =548
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5496
                            Top =312
                            Width =2925
                            Height =240
                            RightMargin =29
                            Name ="labRemoveFilterCtl"
                            Caption ="CommandButton \"Remove Filter\":"
                            Tag ="LANG:"
                            LayoutCachedLeft =5496
                            LayoutCachedTop =312
                            LayoutCachedWidth =8421
                            LayoutCachedHeight =552
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =8423
                    Top =24
                    Width =1659
                    TabIndex =1
                    Name ="cbxApplyFilterCtl"
                    RowSourceType ="Value List"
                    ColumnWidths ="2268"

                    LayoutCachedLeft =8423
                    LayoutCachedTop =24
                    LayoutCachedWidth =10082
                    LayoutCachedHeight =264
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5496
                            Top =28
                            Width =2925
                            Height =240
                            RightMargin =29
                            Name ="labUseFilterCtl"
                            Caption ="CommandButton \"Apply Filter\":"
                            Tag ="LANG:"
                            LayoutCachedLeft =5496
                            LayoutCachedTop =28
                            LayoutCachedWidth =8421
                            LayoutCachedHeight =268
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =2268
                    Left =8424
                    Top =598
                    Width =1659
                    TabIndex =6
                    Name ="cbxAutoFilterCtl"
                    RowSourceType ="Value List"

                    LayoutCachedLeft =8424
                    LayoutCachedTop =598
                    LayoutCachedWidth =10083
                    LayoutCachedHeight =838
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5497
                            Top =602
                            Width =2925
                            Height =240
                            RightMargin =29
                            Name ="labAutoFilterCtl"
                            Caption ="AutoFilter Checkbox:"
                            Tag ="LANG:"
                            LayoutCachedLeft =5497
                            LayoutCachedTop =602
                            LayoutCachedWidth =8422
                            LayoutCachedHeight =842
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =56
                    Top =6122
                    Width =2568
                    Height =240
                    ForeColor =16737792
                    Name ="labCheckVersion"
                    Caption ="Check wizard version"
                    OnClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    Tag ="unchecked|LANG:"
                    VerticalAnchor =1
                    LayoutCachedLeft =56
                    LayoutCachedTop =6122
                    LayoutCachedWidth =2624
                    LayoutCachedHeight =6362
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =87
                    Left =2004
                    Top =595
                    Width =3402
                    Height =284
                    FontSize =7
                    TabIndex =5
                    Name ="cmdFillFilterControlsFromDataSource"
                    Caption ="Fill list with names from data source"
                    StatusBarText ="Please select the data form in the 'ApplyFilter method' option"
                    OnClick ="[Event Procedure]"
                    Tag ="LANG:"
                    ControlTipText ="Please select the data form in the 'ApplyFilter method' option"

                    LayoutCachedLeft =2004
                    LayoutCachedTop =595
                    LayoutCachedWidth =5406
                    LayoutCachedHeight =879
                    BackColor =-2147483633
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10128
                    Top =24
                    Width =561
                    Height =228
                    TabIndex =2
                    BackColor =14151142
                    Name ="cbxLangCode"
                    RowSourceType ="Value List"
                    RowSource ="\"DE\";\"EN\""
                    AfterUpdate ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =10128
                    LayoutCachedTop =24
                    LayoutCachedWidth =10689
                    LayoutCachedHeight =252
                End
                Begin CheckBox
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =7605
                    Top =5191
                    TabIndex =14
                    Name ="cbUseFilterControlTagConverter"
                    DefaultValue ="False"

                    LayoutCachedLeft =7605
                    LayoutCachedTop =5191
                    LayoutCachedWidth =7865
                    LayoutCachedHeight =5431
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7833
                            Top =5160
                            Width =2760
                            Height =240
                            Name ="Label132"
                            Caption ="Filter definition in tag property"
                            Tag ="LANG:"
                            LayoutCachedLeft =7833
                            LayoutCachedTop =5160
                            LayoutCachedWidth =10593
                            LayoutCachedHeight =5400
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =2345
                    Top =5019
                    Width =295
                    Height =295
                    FontSize =16
                    ForeColor =5026082
                    Name ="labCopyModulFilterControltagConverter"
                    FontName ="Marlett"
                    VerticalAnchor =1
                    LayoutCachedLeft =2345
                    LayoutCachedTop =5019
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5314
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =315
                    Top =5044
                    Width =1991
                    Height =255
                    TopMargin =14
                    Name ="labCaptionCopyModulFilterControlTagConverter"
                    Caption ="FilterControlTagConverter"
                    VerticalAnchor =1
                    LayoutCachedLeft =315
                    LayoutCachedTop =5044
                    LayoutCachedWidth =2306
                    LayoutCachedHeight =5299
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =215
                    Left =2345
                    Top =5019
                    Width =295
                    Height =295
                    TabIndex =13
                    Name ="cmdCopyModulFilterControlTagConverter"
                    Caption ="Befehl69"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =2345
                    LayoutCachedTop =5019
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =5314
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "frmFilterFormWizard.cls"
