Operation =6
Option =0
Begin InputTables
    Name ="L10n_Dict"
End
Begin OutputColumns
    Expression ="L10n_Dict.KeyText"
    GroupLevel =2
    Expression ="L10n_Dict.LangCode"
    GroupLevel =1
    Alias ="MaxOfLngText"
    Expression ="Max(L10n_Dict.LngText)"
End
Begin Groups
    Expression ="L10n_Dict.KeyText"
    GroupLevel =2
    Expression ="L10n_Dict.LangCode"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="DE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5565"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="EN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5190"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="EN.KeyText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EN.LangCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DE.LangCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DE.KeyText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="L10n_Dict.LangCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="L10n_Dict.KeyText"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
        dbInteger "ColumnWidth" ="5985"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="L10n_Dict.LngText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1595
    Bottom =836
    Left =-1
    Top =-1
    Right =1571
    Bottom =309
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =61
        Top =66
        Right =205
        Bottom =210
        Top =0
        Name ="L10n_Dict"
        Name =""
    End
End
