// Power Query extracted from: complex.xlsm
// Location: customXml/item1.xml (DataMashup format)
// Extracted on: 2025-07-11T14:27:48.261Z

section Section1;

shared fGetNamedRange = let GetNamedRange=(NamedRange) =>
 
let
	name = Excel.CurrentWorkbook(){[Name=NamedRange]}[Content],
	value = name{0}[Column1]
in
	value

in GetNamedRange;

shared RawInput = let
    Source = fGetNamedRange("InputText"),
    #"Converted to Table" = #table(1, {{Source}}),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "RawInput"}})
in
    #"Renamed Columns";

shared FinalTable = let
    Raw = RawInput,
    AddedDate = Table.AddColumn(Raw, "Now", each DateTime.LocalNow())
in
    AddedDate
;