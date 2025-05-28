// ---------------------------------------------
// World Bank Indicator Processor
// Author: Aleksei Ogarkov
// Date: 2025-04-30
// Description:
// Enhances raw World Bank data with country and indicator metadata,
// applies value scaling (e.g., GDP → billions, Population → millions),
// and prepares a clean, structured dataset for reporting and analysis.
//
// Output: Enriched and normalized indicator dataset with consistent units.
// ---------------------------------------------

let
    // === PARAMETERS ===
    // Load base data and parameter tables from Excel
    source           = RawDataQuery_vVBA,       // alternative options: RawDataQuery_vPQ / RawDataQuery_vVBA
    countryTable     = Excel.CurrentWorkbook(){[Name = "Countries"]}[Content],
    indicatorTable   = Excel.CurrentWorkbook(){[Name = "Indicators"]}[Content],
    startYear        = Excel.CurrentWorkbook(){[Name = "StartYear"]}[Content]{0}[Column1],
    endYear          = Excel.CurrentWorkbook(){[Name = "EndYear"]}[Content]{0}[Column1],
    countries        = countryTable[CountryCode],
    indicators       = indicatorTable[IndicatorCode],

    // === TRANSFORMATION ===
    // Round and cast column types for consistent processing
    TransformedSource = Table.TransformColumnTypes(
        Table.TransformColumns(source, {
            {"Value", each Number.Round(_, 2), type number}
        }),
        {
            {"Year", Int64.Type},
            {"Indicator", type text},
            {"Country", type text}
        }
    ),

    // === METADATA JOINING ===
    // Merge country metadata
    MergedCountries = Table.NestedJoin(
        TransformedSource,
        {"Country"},
        DimCountries,
        {"cCode"},
        "CountryData",
        JoinKind.LeftOuter
    ),
    ExpandedCountries = Table.ExpandTableColumn(
        MergedCountries,
        "CountryData",
        {"Country", "Region", "Weight", "Income Group"},
        {"Countries.Country", "Countries.Region", "Countries.Weight", "Countries.Income Group"}
    ),

    // Merge indicator metadata
    MergedIndicators = Table.NestedJoin(
        ExpandedCountries,
        {"Indicator"},
        DimIndicators,
        {"IndicatorCode"},
        "Indicators",
        JoinKind.LeftOuter
    ),
    ExpandedIndicators = Table.Buffer(Table.ExpandTableColumn(
        MergedIndicators,
        "Indicators",
        {"Indicator", "Group", "Unit", "Type"},
        {"Indicators.Indicator", "Indicators.Group", "Indicators.Unit", "Indicators.Type"}
    )),

    // === SCALING LOGIC ===
    // Define transformation rules for scaling by indicator type
    WithScaleFactor = Table.AddColumn(ExpandedIndicators, "ScaleFactor", each
        let indicator = [Indicators.Indicator]
        in if List.Contains({"GDP", "High-tech exports", "ICT service exports"}, indicator) then 1e9
           else if indicator = "Population" then 1e3
           else null,
        type nullable number
    ),

    ScaledValues = Table.AddColumn(WithScaleFactor, "AdjustedValue", each
        if [ScaleFactor] <> null then [Value] / [ScaleFactor] else [Value],
        type number
    ),

    ScaledUnits = Table.AddColumn(ScaledValues, "AdjustedUnit", each
        let indicator = [Indicators.Indicator]
        in if List.Contains({"GDP", "High-tech exports", "ICT service exports"}, indicator) then "B USD"
           else if indicator = "Population" then "M"
           else [Indicators.Unit],
        type text
    ),

    // === FINAL CLEANUP ===
    Cleaned = Table.RemoveColumns(ScaledUnits, {"Value", "Indicators.Unit", "ScaleFactor"}),
    Renamed = Table.RenameColumns(Cleaned, {
        {"AdjustedValue", "Value"},
        {"AdjustedUnit", "Indicators.Unit"}
    }),
    Final = Table.ReorderColumns(Renamed, Table.ColumnNames(ExpandedIndicators))
in
    Final