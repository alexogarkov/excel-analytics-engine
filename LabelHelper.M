// ---------------------------------------------
// Weighted Regional Aggregation
// Author: Aleksei Ogarkov
// Date: 2025-05-27
//
// Description:
// Performs region-level weighted aggregation on a filtered set of macroeconomic indicators.
// Differentiates between absolute and relative indicators,
// applies appropriate weighting logic, and computes normalized results.
// 
// Output Columns:
// GeoName | Income Group | Indicator | Year | Value | ValueType | Weight
// ---------------------------------------------

let
    Source = Table.Buffer(Dataset),

    // Filter only relevant indicators for aggregation
    FilteredIndicators = Table.SelectRows(Source, each List.Contains(
        {
            "Fixed broadband subscriptions", 
            "GDP growth", 
            "GDP per capita", 
            "Internet users", 
            "Labor force with advanced edu", 
            "Unemployment rate"
        }, 
        [Indicators.Indicator]
    )),

    // Assign region as aggregation unit
    ExpandedRegion = Table.AddColumn(FilteredIndicators, "AggGeoList", each [Countries.Region]),

    // Add consistent geo structure metadata
    WithGeoStructure = Table.AddColumn(ExpandedRegion, "GeoLevel", each "Region", type text),
    WithGeoName = Table.AddColumn(WithGeoStructure, "GeoName", each [AggGeoList], type text),

    // Compute weighted values:
    // - Absolute (Type = "A") → use raw value
    // - Relative (Type = "R") → multiply by country weight
    AddWeightedData = Table.AddColumn(WithGeoName, "WeightedValue", each 
        if [Indicators.Type] = "A" then [Value] 
        else [Value] * [Countries.Weight], 
        type nullable number
    ),

    // Include weights only for relative indicators
    AddWeightColumn = Table.AddColumn(AddWeightedData, "IncludedWeight", each 
        if [Indicators.Type] = "A" then null 
        else [Countries.Weight], 
        type nullable number
    ),

    // Group data by region, indicator, year, and type; then aggregate weighted values and weights
    Grouped = Table.Group(AddWeightColumn, 
        {"GeoLevel", "GeoName", "Indicators.Indicator", "Year", "Indicators.Type", "Countries.Income Group"}, 
        {
            {"WeightedSum", each List.Sum(List.RemoveNulls([WeightedValue])), type nullable number},
            {"WeightSum", each List.Sum(List.RemoveNulls([IncludedWeight])), type nullable number}
        }
    ),

    // Calculate final indicator value:
    // - Sum for absolute indicators
    // - Weighted average for relative indicators
    AddFinalValue = Table.AddColumn(Grouped, "Value", each 
        if [Indicators.Type] = "A" then [WeightedSum]
        else if ([WeightSum] = null or [WeightSum] = 0) then null
        else [WeightedSum] / [WeightSum], 
        type nullable number
    ),

    // Rename and select final output columns
    FinalTable = Table.SelectColumns(
        Table.RenameColumns(AddFinalValue, {
            {"Countries.Income Group", "Income Group"},
            {"Indicators.Indicator", "Indicator"},
            {"Indicators.Type", "ValueType"},
            {"WeightSum", "Weight"}
        }),
        {"GeoName", "Income Group", "Indicator", "Year", "Value", "ValueType", "Weight"}
    ),

    // Round final values to 4 decimal places
    Final = Table.TransformColumns(FinalTable, {
        {"Value", each Number.Round(Number.FromText(Text.From(_)), 4), type number}
    })
in
    Final
