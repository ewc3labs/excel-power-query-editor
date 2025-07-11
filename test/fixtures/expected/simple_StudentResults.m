// Power Query extracted from: simple.xlsx
// Location: customXml/item1.xml (DataMashup format)
// Extracted on: 2025-06-22T04:07:14.577Z

section Section1;

shared StudentResults = let
    Source = Excel.CurrentWorkbook(){[Name="StudentNames"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Name", type text}, {"Age", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "DateTimeGenerated", each DateTime.LocalNow())
in
    #"Added Custom";
