// ---------------------------------------------
// Latest Available Indicators
// Author: Aleksei Ogarkov
// Date: 2025-04-30
// Description:
// Groups the dataset by Country and Indicator,
// extracts values for the two latest years per group,
// reshapes into a flat structure with columns:
// Region | Country | Indicator | Type | Weight | ValueLastYear | ValuePreviousYear
// ---------------------------------------------

let
    // === STEP 1: LOAD SOURCE DATA ===
    Source = Dataset,

    // === STEP 2: GROUP AND SORT ===
    GroupedRows = Table.Group(
        Source,
        {"Countries.Country", "Countries.Region", "Countries.Income Group", "Countries.Weight", "Indicators.Indicator", "Indicators.Unit", "Indicators.Type"},
        {
            {
                "GroupedData",
                each Table.FirstN(
                    Table.Sort(_, {{"Year", Order.Descending}}),
                    2
                ),
                type table [
                    Countries.Region = nullable text,
                    Countries.Country = text,
                    Countries.Income group = text,
                    Indicators.Indicator = text,
                    Indicators.Type = text,
                    Indicators.Unit = text,
                    Countries.Weight = nullable number,
                    Year = nullable number,
                    Value = nullable number
                ]
            }
        }
    ),

    // === STEP 3: EXTRACT STRUCTURED RECORDS ===
    Extracted = Table.AddColumn(
        GroupedRows,
        "ExtractedValues",
        each [
            Region = Record.FieldOrDefault([GroupedData]{0}, "Countries.Region"),
            Country = Record.FieldOrDefault([GroupedData]{0}, "Countries.Country"),
            Income group = Record.FieldOrDefault([GroupedData]{0}, "Countries.Income Group"),
            Weight = Record.FieldOrDefault([GroupedData]{0}, "Countries.Weight"),
            Indicator = Record.FieldOrDefault([GroupedData]{0}, "Indicators.Indicator"),
            Unit = Record.FieldOrDefault([GroupedData]{0}, "Indicators.Unit"),
            Type = Record.FieldOrDefault([GroupedData]{0}, "Indicators.Type"),
            ValueLastYear = try Number.From(Record.FieldOrDefault([GroupedData]{0}, "Value")) otherwise null,
            ValuePreviousYear = try Number.From(Record.FieldOrDefault([GroupedData]{1}, "Value")) otherwise null
        ]
    ),

    // === STEP 4: BUILD FINAL TABLE WITH EXPLICIT TYPE DEFINITION ===
    FinalTable = Table.FromRecords(Extracted[ExtractedValues]),

    Final = Table.TransformColumns(FinalTable, {
        {"Weight", each _, type nullable number},
        {"ValueLastYear", each Number.Round(Number.FromText(Text.From(_)), 4), type number},
        {"ValuePreviousYear", each Number.Round(Number.FromText(Text.From(_)), 4), type number}
    }),

    #"Filtered Rows" = Table.SelectRows(Final, each true)
in
    #"Filtered Rows"