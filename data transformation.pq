let
    PATH = Main_Folder_Path,

    Folder =  
        let
            FILES = Folder.Contents(PATH) 
        in 
            let 
                GetExcelContent = 
                    Table.TransformColumns(FILES, {"Content", each 
                        let 
                            content = _ 
                        in 
                            Excel.Workbook(content) })
            in 
                Table.RenameColumns(GetExcelContent, {"Content", "Data"}),

    #"Extraction: Filtered Out (~)" = 
        let 
            TABLE = Folder
        in 
            Table.SelectRows(TABLE, each 
                not Text.StartsWith([Name], "~")),
    
    #"Extraction: Get latest File" = 
        let 
            TABLE = #"Extraction: Filtered Out (~)"
        in
            Table.SelectRows(TABLE, 
                let 
                    LATEST = List.Max(TABLE[Date modified]) 
                in 
                    each [Date modified] = LATEST),

    #"Extraction: Removed Other Columns" =     
        let
            TABLE = #"Extraction: Get latest File"
        in
            Table.SelectColumns(TABLE,{"Data"}),

    #"Extraction: Filtred Out (_)" = 
        let 
            TABLE = #"Extraction: Removed Other Columns"
        
        in 
            Table.TransformColumns(TABLE, {"Data", each 
                let 
                    INNER_TABLE = _ 
                in 
                    Table.SelectRows(INNER_TABLE, each 
                        not Text.StartsWith([Name], "_") ) } ),
    
    #"Extraction: Expanded Data" = 
        let 
            TABLE = #"Extraction: Filtred Out (_)",
            COLUMN_LIST = {"Name", "Data"}
            // COLUMN_LIST = {"Data"}
        in 
            Table.ExpandTableColumn(TABLE, "Data",COLUMN_LIST, COLUMN_LIST),
    
    #"Extraction: Repalce Values" = 
        let 
            TABLE = #"Extraction: Expanded Data",
            OLD_VALUE = "CR-",
            NEW_VALUE = "",
            COLUMN = "Name"
        in
            Table.ReplaceValue(TABLE ,OLD_VALUE,NEW_VALUE,Replacer.ReplaceText,{COLUMN}),

    #"Extraction: Filter Sheet"  = 
    
        let 
            TABLE = #"Extraction: Repalce Values"
        in 
            Table.SelectRows(TABLE, each 
                let 
                    NUM = Text.Select( [Name], {"0".."9"} ),
                    YEAR = 
                        if Text.Length(NUM) = 2
                        then 2000 + Number.From(UM)
                        else Number.From(NUM),
                    APPLIED_FILTER = 2030 >= YEAR and YEAR >= 2024 and not Text.StartsWith([Name], "_")    
                in 
                    APPLIED_FILTER),

    #"Extraction: Navigate to Latest Sheet" = 
        let 
            TABLE = #"Extraction: Filter Sheet", 
            
            CONVERT_TO_DATE = List.Transform(TABLE[Name], each 
                let 
                    DATE = _ 
                in 
                    Date.From(DATE)), 
            LATEST = List.Max(CONVERT_TO_DATE),

            SELECT_ROW = Table.SelectRows(TABLE, each 
                let 
                    DATE = Date.From([Name])
                in 
                    DATE = LATEST),
            COLUMN = {"Data"}
        in
            Table.SelectColumns(SELECT_ROW, COLUMN),

    #"Transformation: Remove Top Rows" = 
        let 
            TABLE = #"Extraction: Navigate to Latest Sheet"
        in
            Table.TransformColumns(TABLE, {"Data", each 
                let 
                        INNER_TABLE = _,
                        COLUMN_NAME_LIST = Table.ColumnNames(INNER_TABLE),
                        FIRST_COLUMN_NAME = List.First(COLUMN_NAME_LIST),
                        FIRST_COLUMN_VALUE = Table.Column(INNER_TABLE, FIRST_COLUMN_NAME),
                        MATCHES = List.FindText(FIRST_COLUMN_VALUE, "No."),
                        GET_POSITION_OF_SAP = List.PositionOf(FIRST_COLUMN_VALUE, MATCHES{0})
                    in  
                        
                        Table.Skip(INNER_TABLE, GET_POSITION_OF_SAP)
            }),

    #"Transformation: Promote Header" = 
        let
            TABLE = #"Transformation: Remove Top Rows"
        in
            Table.TransformColumns(TABLE, {"Data", each 
            let
                INNER_TABLE = _
            in
                Table.PromoteHeaders(INNER_TABLE)
            }),

    
    #"Transformation: Table Name Header List" = 
        let
            TABLE = #"Transformation: Promote Header"
        in
            Table.FromList(Table.ColumnNames(TABLE{0}[Data])),

    #"**Transformation: Filtered Rows" = Table.SelectRows(#"Transformation: Table Name Header List", each 
        (
            [Column1] = "Estimated Amount / Final Claimed (OMR)" or 
            [Column1] = "Item" or 
            [Column1] = "Transmittals " or 
            [Column1] = "Purchase Order (PO)" or 
            [Column1] = "Verified Estimates / Certified Amount (OMR)" or 
            [Column1] = "Verifier"
            )
    ),
    
    #"Transformation: Keep Selected Columns" = 
        let
            TABLE = #"Transformation: Promote Header"
        in
            Table.TransformColumns( TABLE, {"Data", each
                let 
                    innerTable = _,
                    ColumnNameList = Table.ColumnNames(innerTable)
                in 
                    Table.SelectColumns(_, #"**Transformation: Filtered Rows"[Column1])
            }),

    #"**Transformation: Select Transmittal" = 
        Table.SelectRows(#"Transformation: Table Name Header List", each 
            ([Column1] = "Transmittals ")),

    #"Transformation: Filter PO No" = 
        let 
            TABLE = #"Transformation: Keep Selected Columns"
        in
            Table.TransformColumns(TABLE, {"Data", each 
                let 
                    INNER_TABLE = _,
                    RENAMED_COLUMN = Table.RenameColumns(INNER_TABLE, {List.First( Table.ColumnNames(INNER_TABLE) ),"PO"}),
                    RENAMED_TRANSMITTAL =  Table.RenameColumns(RENAMED_COLUMN, { List.First(#"**Transformation: Select Transmittal"[Column1]), "Transmittal"}  )
                in 
                    Table.SelectRows(RENAMED_TRANSMITTAL, each 
                        Text.StartsWith(Text.From([PO]), "4500"))
            }),

    #"**Transformation: Transmittal Title" = #"Transmittal Generator"(5, -4),

    #"Transformation: Filter Trasmittal" = 
        let 
            TABLE = #"Transformation: Filter PO No"
        in
            Table.TransformColumns( TABLE,{"Data", each 
                let
                    INNER_TABLE = _ 
                in
                    Table.SelectRows(INNER_TABLE, each 
                        let
                            TRASMITTAL_NO = Number.From(Text.Select([Transmittal], {"0".."9"}))
                        in
                            List.Contains(#"**Transformation: Transmittal Title"[Transmittal No], Text.From(TRASMITTAL_NO) )
                    )
            }),

    #"Transformation: Add Date Column" = 
    let
        TABLE = #"Transformation: Filter Trasmittal"
    in
        Table.TransformColumns( TABLE, {"Data", each 
            let 
                innerTable = _
            in  
                Table.AddColumn(innerTable, "Date", each 
                    let 
                        TRASMITTAL_NO = Number.From( Text.Select( Text.From([Transmittal]), {"0".."9"} ) ),
                        GetPostionOfTransmittal =  List.PositionOf(#"**Transformation: Transmittal Title"[#"Transmittal No"], Text.From(TRASMITTAL_NO) ),
                        Date = #"**Transformation: Transmittal Title"[#"Date"]{ GetPostionOfTransmittal }
                    in
                        Date 
                    
                ) 
        }),

    #"Transformation: Add Contract" = 
        let
            TABLE = #"Transformation: Add Date Column"
        in
            Table.TransformColumns(TABLE, {"Data", each 
                let 
                    INNER_TABLE = _,
                    NEW_COLUMN = "Contract",
                    OFF_SET = Text.PositionOf(PATH, "C31"), 
                    TEXT = Text.Range(PATH,OFF_SET),
                    CONTRACT_NO = 
                        if Text.Contains(TEXT, "–") then 
                            Text.BeforeDelimiter(TEXT, "–") 
                        else  
                            Text.BeforeDelimiter(TEXT, "-")
                in
                Table.AddColumn(INNER_TABLE, NEW_COLUMN, each CONTRACT_NO) 
            }),

    #"Transformation: Add Location" = 
        let
            TABLE = #"Transformation: Add Contract"
        in
            Table.TransformColumns(TABLE, {"Data", each 
                let
                    INNER_TABLE = _,
                    NEW_COLUMN = "Location",
                    LOCATION_VALUE = "Fahud"
                in
                    Table.AddColumn(INNER_TABLE, NEW_COLUMN, each LOCATION_VALUE)       
            }),

    #"Transformation: Remove Transmittal Column" = 
        let
            TABLE = #"Transformation: Add Location"
        in
            Table.TransformColumns(TABLE, {"Data", each 
                let
                    INNER_TABLE = _,
                    COLUMN_TO_REMOVE = "Transmittal"
                in
                    Table.RemoveColumns(INNER_TABLE, COLUMN_TO_REMOVE )        
            }),

    #"Transformation: Merage Columns" = 
        let
            TABLE = #"Transformation: Remove Transmittal Column"
        in
            Table.TransformColumns( TABLE, {"Data", each 
            
                let 
                    INNER_TABLE = _,
                    COLUMN_NAME_LIST= Table.ColumnNames(INNER_TABLE)
                in 
                    Table.CombineColumns(
                        Table.TransformColumnTypes(INNER_TABLE, {{COLUMN_NAME_LIST{0}, type text}, {COLUMN_NAME_LIST{1}, type text}}, "en-US"),{COLUMN_NAME_LIST{0}, COLUMN_NAME_LIST{1}},
                        Combiner.CombineTextByDelimiter("/", QuoteStyle.None),COLUMN_NAME_LIST{0})
            }),

    #"Transformation: Rename Columns" = 
        
        let
            TABLE = #"Transformation: Merage Columns"
        in
            Table.TransformColumns( TABLE, {"Data", each 
                let 
                    Header = Table.ColumnNames(_) 
                in 
                    Table.RenameColumns(_ ,{
                        { Header{0}, #"Column Name List"{0} }, 
                        { Header{1}, #"Column Name List"{1} }, 
                        { Header{2}, #"Column Name List"{2} }, 
                        { Header{3}, #"Column Name List"{3} }, 
                        { Header{4}, #"Column Name List"{4} },
                        { Header{5}, #"Column Name List"{5} } 
                    })
            })
in
    #"Transformation: Rename Columns"