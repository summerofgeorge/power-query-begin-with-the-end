let
    // Shared preparation function, applied once per regional source table
    PrepareRegion = (SourceTable as table, RegionName as text) as table =>
        let
            FilledBranch = Table.FillDown(SourceTable, {"Branch"}),
            DetailRowsOnly = Table.SelectRows(FilledBranch, each [RowType] = "Detail"),
            DroppedRowType = Table.RemoveColumns(DetailRowsOnly, {"RowType"}),
            SplitSkuProduct = Table.SplitColumn(
                DroppedRowType, "Product",
                Splitter.SplitTextByDelimiter(" | ", QuoteStyle.None), {"SKU", "Product"}),
            TextSkuProduct = Table.TransformColumnTypes(SplitSkuProduct, {{"SKU", type text}, {"Product", type text}}),
            // Every column other than these three is treated as a date column, so new ISO date columns are included
            UnpivotedDates = Table.UnpivotOtherColumns(TextSkuProduct, {"Branch", "SKU", "Product"}, "Date", "Sales"),
            // Blank (null) amounts are "not reported" and are omitted. Zeros and text such as " " are kept
            ReportedOnly = Table.SelectRows(UnpivotedDates, each [Sales] <> null),
            AddedRegion = Table.AddColumn(ReportedOnly, "Region", each RegionName, type text),
            // No try/otherwise, so conversion errors stay visible in the output
            ConvertedTypes = Table.TransformColumns(AddedRegion, {
                {"Date", each Date.FromText(_, [Format = "yyyy-MM-dd", Culture = "en-US"]), type date},
                {"Sales", each if Value.Is(_, type number) then _ else Number.FromText(_, "en-US"), Currency.Type}
            })
        in
            ConvertedTypes,

    EastPrepared = PrepareRegion(Excel.CurrentWorkbook(){[Name = "Raw_East"]}[Content], "East"),
    WestPrepared = PrepareRegion(Excel.CurrentWorkbook(){[Name = "Raw_West"]}[Content], "West"),
    CombinedRegions = Table.Combine({EastPrepared, WestPrepared}),
    OrderedColumns = Table.SelectColumns(CombinedRegions, {"Region", "Branch", "SKU", "Product", "Date", "Sales"})
in
    OrderedColumns
