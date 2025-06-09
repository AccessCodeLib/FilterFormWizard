Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10050
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =3300
    Top =3443
    Right =13635
    Bottom =6428
    RecSrcDt = Begin
        0xc4d680b8ef58e440
    End
    RecordSource ="tabFilterControls"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            ForeThemeColorIndex =2
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
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
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =285
            Name ="Formularkopf"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =5831
                    Width =2040
                    Height =285
                    FontSize =10
                    Name ="labControl"
                    Caption ="Control"
                    FontName ="Tahoma"
                    Tag ="LANG:"
                    LayoutCachedLeft =5831
                    LayoutCachedWidth =7871
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =56
                    Width =2157
                    Height =285
                    FontSize =10
                    Name ="labDataField"
                    Caption ="Data field"
                    FontName ="Tahoma"
                    Tag ="LANG:"
                    LayoutCachedLeft =56
                    LayoutCachedWidth =2213
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3746
                    Width =1815
                    Height =285
                    FontSize =10
                    Name ="labRelaionalOperator"
                    Caption ="Relational operator"
                    FontName ="Tahoma"
                    Tag ="LANG:"
                    LayoutCachedLeft =3746
                    LayoutCachedWidth =5561
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =2268
                    Width =1380
                    Height =285
                    FontSize =10
                    Name ="labDataType"
                    Caption ="Data type"
                    FontName ="Tahoma"
                    Tag ="LANG:"
                    LayoutCachedLeft =2268
                    LayoutCachedWidth =3648
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =7875
                    Width =2175
                    Height =285
                    FontSize =10
                    Name ="labControl2"
                    Caption ="Other controls"
                    FontName ="Tahoma"
                    Tag ="LANG:"
                    LayoutCachedLeft =7875
                    LayoutCachedWidth =10050
                    LayoutCachedHeight =285
                    ThemeFontIndex =-1
                End
            End
        End
        Begin Section
            Height =642
            Name ="Detailbereich"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =97.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =3402
                    Left =5831
                    Top =29
                    Width =2041
                    Height =285
                    FontSize =10
                    TabIndex =3
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000000000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b006300620043007200650061007400650043006f006e00740072006f006c00 ,
                        0x5d003d00540072007500650000000000
                    End
                    Name ="fcControl"
                    ControlSource ="Control"
                    RowSourceType ="Value List"
                    ColumnWidths ="3402"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5831
                    LayoutCachedTop =29
                    LayoutCachedWidth =7872
                    LayoutCachedHeight =314
                    ThemeFontIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000000000000000000ffffff00160000005b00 ,
                        0x6300620043007200650061007400650043006f006e00740072006f006c005d00 ,
                        0x3d005400720075006500000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3746
                    Top =29
                    Width =1803
                    Height =285
                    FontSize =10
                    TabIndex =2
                    BoundColumn =1
                    Name ="fcRelationalOperator"
                    ControlSource ="RelationalOperator"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tabRelationalOperators.RelationalOperator, tabRelationalOperators.Relatio"
                        "nalOperatorCode FROM tabRelationalOperators ORDER BY tabRelationalOperators.Orde"
                        "rPos;"
                    ColumnWidths ="1134;0"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3746
                    LayoutCachedTop =29
                    LayoutCachedWidth =5549
                    LayoutCachedHeight =314
                    ThemeFontIndex =-1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2268
                    Top =29
                    Width =1353
                    Height =285
                    FontSize =10
                    TabIndex =1
                    BoundColumn =1
                    Name ="fcDataType"
                    ControlSource ="DataType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tabSqlFieldDataTypes.SqlFieldDataType, tabSqlFieldDataTypes.SqlFieldDataT"
                        "ypeCode FROM tabSqlFieldDataTypes ORDER BY tabSqlFieldDataTypes.SqlFieldDataType"
                        ";"
                    ColumnWidths ="1134;0"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2268
                    LayoutCachedTop =29
                    LayoutCachedWidth =3621
                    LayoutCachedHeight =314
                    ThemeFontIndex =-1
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5465
                    Top =390
                    TabIndex =7
                    Name ="fcRelationalOperatorNot"
                    ControlSource ="RelationalOperatorNot"
                    DefaultValue ="False"

                    LayoutCachedLeft =5465
                    LayoutCachedTop =390
                    LayoutCachedWidth =5725
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =127
                            Left =5135
                            Top =360
                            Width =330
                            Height =240
                            FontSize =8
                            Name ="labRelationalOperatorNot"
                            Caption ="Not"
                            FontName ="Tahoma"
                            LayoutCachedLeft =5135
                            LayoutCachedTop =360
                            LayoutCachedWidth =5465
                            LayoutCachedHeight =600
                            ThemeFontIndex =-1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4181
                    Top =390
                    TabIndex =5
                    Name ="fcWildCardSuffix"
                    ControlSource ="WildCardSuffix"
                    DefaultValue ="False"

                    LayoutCachedLeft =4181
                    LayoutCachedTop =390
                    LayoutCachedWidth =4441
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =3746
                            Top =360
                            Width =435
                            Height =240
                            FontSize =8
                            Name ="labWildCardSuffix"
                            Caption ="xxx*"
                            FontName ="Tahoma"
                            LayoutCachedLeft =3746
                            LayoutCachedTop =360
                            LayoutCachedWidth =4181
                            LayoutCachedHeight =600
                            ThemeFontIndex =-1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =4886
                    Top =390
                    TabIndex =6
                    Name ="fcWildCardPrefix"
                    ControlSource ="WildCardPrefix"
                    DefaultValue ="False"

                    LayoutCachedLeft =4886
                    LayoutCachedTop =390
                    LayoutCachedWidth =5146
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =4451
                            Top =360
                            Width =435
                            Height =240
                            FontSize =8
                            Name ="labWildCardPrefix"
                            Caption ="*xxx"
                            FontName ="Tahoma"
                            LayoutCachedLeft =4451
                            LayoutCachedTop =360
                            LayoutCachedWidth =4886
                            LayoutCachedHeight =600
                            ThemeFontIndex =-1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =56
                    Top =29
                    Width =2160
                    Height =285
                    FontSize =10
                    Name ="fcDataField"
                    ControlSource ="DataField"
                    FontName ="Tahoma"

                    LayoutCachedLeft =56
                    LayoutCachedTop =29
                    LayoutCachedWidth =2216
                    LayoutCachedHeight =314
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7938
                    Top =29
                    Width =2056
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="fcControl2"
                    ControlSource ="Control2"
                    StatusBarText ="Mehrere Steuerelemente mit Komma trennen"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000000000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b006300620043007200650061007400650043006f006e00740072006f006c00 ,
                        0x5d003d00540072007500650000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedLeft =7938
                    LayoutCachedTop =29
                    LayoutCachedWidth =9994
                    LayoutCachedHeight =314
                    ThemeFontIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000000000000000000ffffff00160000005b00 ,
                        0x6300620043007200650061007400650043006f006e00740072006f006c005d00 ,
                        0x3d005400720075006500000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6222
                    Top =390
                    TabIndex =8
                    Name ="cbCreateControl"
                    ControlSource ="CreateControl"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"

                    LayoutCachedLeft =6222
                    LayoutCachedTop =390
                    LayoutCachedWidth =6482
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =5831
                            Top =360
                            Width =414
                            Height =240
                            FontSize =8
                            Name ="labCreateControl"
                            Caption ="New"
                            FontName ="Tahoma"
                            Tag ="LANG:"
                            LayoutCachedLeft =5831
                            LayoutCachedTop =360
                            LayoutCachedWidth =6245
                            LayoutCachedHeight =600
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =0
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionGroup
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =5839
                    Top =336
                    Width =4137
                    Height =300
                    TabIndex =9
                    Name ="ogNewControlType"
                    ControlSource ="ControlType"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedLeft =5839
                    LayoutCachedTop =336
                    LayoutCachedWidth =9976
                    LayoutCachedHeight =636
                    Begin
                        Begin OptionButton
                            OverlapFlags =119
                            Left =6689
                            Top =396
                            OptionValue =109
                            Name ="ogfldTextbox"

                            LayoutCachedLeft =6689
                            LayoutCachedTop =396
                            LayoutCachedWidth =6949
                            LayoutCachedHeight =636
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6919
                                    Top =336
                                    Width =876
                                    Height =300
                                    Name ="labCreateTextbox"
                                    Caption ="TextBox"
                                    LayoutCachedLeft =6919
                                    LayoutCachedTop =336
                                    LayoutCachedWidth =7795
                                    LayoutCachedHeight =636
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =7992
                            Top =396
                            TabIndex =1
                            OptionValue =111
                            Name ="ogfldCombobox"

                            LayoutCachedLeft =7992
                            LayoutCachedTop =396
                            LayoutCachedWidth =8252
                            LayoutCachedHeight =636
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8219
                                    Top =336
                                    Width =1032
                                    Height =300
                                    Name ="labCreateCombobox"
                                    Caption ="ComboBox"
                                    LayoutCachedLeft =8219
                                    LayoutCachedTop =336
                                    LayoutCachedWidth =9251
                                    LayoutCachedHeight =636
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="Formularfuß"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "frmFilterFormWizard_SF_Controls.cls"
