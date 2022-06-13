With Selection 
    .Collapse Direction:=wdCollapseEnd 
    .Range.InsertDatabase _ 
        Format:=wdTableFormatSimple2, Style:=191, _ 
        LinkToSource:=False, Connection:="Entire Spreadsheet", _ 
        DataSource:="C:\MSOffice\Excel\Data.xls" 
End With