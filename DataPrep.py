from pathlib import Path

script_dir = Path(__file__).resolve().parent

data_folder = script_dir / "Data"
pq_path = str(data_folder)

# Remove any \OneDrive segment from the path, it breaks the tmdl
parts = Path(pq_path).parts
parts = [p for p in parts if p.lower() != "onedrive"]

if len(parts) > 1:
    pq_path = str(Path(parts[0], *parts[1:]))
else:
    pq_path = str(Path(parts[0]))

target_file = (
    script_dir
    / "SolArch.SemanticModel"
    / "definition"
    / "tables"
    / "Data.tmdl"
)

target_file.parent.mkdir(parents=True, exist_ok=True)

tmdl_content = r'''
table Data
	lineageTag: d87883eb-f23a-4c18-8a88-16c8004f348d

	measure Notes =
			
			SWITCH (
			    TRUE(),
			    ISINSCOPE('Data'[Concept]),
			        CONCATENATEX(
			            VALUES('Data'[Key Notes]),
			            'Data'[Key Notes],
			            ", "
			        ),
			    ISINSCOPE('Data'[Category]),
			        "...",
			    BLANK()
			)
		lineageTag: e8a66f37-b226-4aad-8b1e-b8d5b6d2476b

	measure Keyword =
			
			SWITCH (
			    TRUE(),
			    ISINSCOPE('Data'[Concept]),
			        CONCATENATEX(
			            VALUES('Data'[Sub-Categories]),
			            'Data'[Sub-Categories],
			            ", "
			        ),
			    ISINSCOPE('Data'[Category]),
			        "...",
			    BLANK()
			)
		lineageTag: 00523223-3cd4-408d-8bbd-ab4cedf24e54

	column 'Source.Name'
		dataType: string
		lineageTag: d4d0ac76-031b-4ca6-b055-ca2e6a37d7f6
		summarizeBy: none
		sourceColumn: Source.Name

		annotation SummarizationSetBy = Automatic

	column Concept
		dataType: string
		lineageTag: 9f47ab6a-f04d-4547-a5f4-be510e82c848
		summarizeBy: none
		sourceColumn: Concept

		annotation SummarizationSetBy = Automatic

	column Category
		dataType: string
		lineageTag: 24d4323b-3887-4b1c-b87d-be43cb517207
		summarizeBy: none
		sourceColumn: Category

		annotation SummarizationSetBy = Automatic

	column 'Key Notes'
		dataType: string
		lineageTag: ad5f890d-be5b-4ca7-9cab-bc843cc39442
		summarizeBy: none
		sourceColumn: Key Notes

		annotation SummarizationSetBy = Automatic

	column Sub-Categories
		dataType: string
		lineageTag: ae505daa-8527-4ed8-ab1f-17e825eb1a01
		summarizeBy: none
		sourceColumn: Sub-Categories

		annotation SummarizationSetBy = Automatic

	column Pillar
		dataType: string
		lineageTag: 8c95eee9-7103-4439-83c9-29948ba5b54a
		summarizeBy: none
		sourceColumn: Pillar

		annotation SummarizationSetBy = Automatic

	column 'Broad Sub-Category'
		dataType: string
		lineageTag: 89d88d29-8cd8-4709-8b6d-e8ba0615cdbc
		summarizeBy: none
		sourceColumn: Broad Sub-Category

		annotation SummarizationSetBy = Automatic

	partition Data = m
		mode: import
		source = ```
				let
				    Source = Folder.Files("SOMEPATH"),
				    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
				    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
				    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
				    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
				    #"Expanded Table Column1" = Table.ExpandTableColumn(
				        #"Removed Other Columns1",
				        "Transform File",
				        Table.ColumnNames(#"Transform File"(#"Sample File"))
				    ),
				    #"Changed Type" = Table.TransformColumnTypes(
				        #"Expanded Table Column1",
				        {
				            {"Source.Name", type text},
				            {"Concept", type text},
				            {"Category", type text},
				            {"Key Notes", type text},
				            {"Pillar", type text}
				        }
				    ),
				
				    MappingTable = #"Sub-Categories",
				
				    #"Mapped SubCategory" = Table.AddColumn(
				        #"Changed Type",
				        "Sub-Categories",
				        each 
				            let
				                notes = Text.Lower([Key Notes]),
				                matches = List.Select(
				                    MappingTable[#"Sub-Category"],
				                    (k) => Text.Contains(notes, Text.Lower(k))
				                ),
				                result = if List.Count(matches) > 0 then Text.Combine(matches, ", ") else null
				            in result,
				        type text
				    ),
				
				    #"Broad Sub-Category" = Table.AddColumn(
				        #"Mapped SubCategory",
				        "Broad Sub-Category",
				        each 
				            let
				                notes = Text.Lower([Key Notes]),
				                mappingList = List.Transform(
				                    MappingTable[#"Sub-Category"],
				                    each Text.Lower(Text.Trim(Text.From(_)))
				                ),
				                firstMatch = List.First(
				                    List.Select(mappingList, (k) => Text.Contains(notes, k)),
				                    null
				                ),
				                result = if firstMatch <> null then
				                    Text.Upper(Text.Start(firstMatch, 1)) & Text.Middle(firstMatch, 1)
				                else
				                    null
				            in result,
				        type text
				    )
				in
				    #"Broad Sub-Category"
				```

	annotation PBI_NavigationStepName = Navigation

	annotation PBI_ResultType = Table

'''

tmdl_content = tmdl_content.replace(
    r'Source = Folder.Files("SOMEPATH"),',
    f'Source = Folder.Files("{pq_path}"),'
)

with open(target_file, "w", encoding="utf-8", newline="\n") as f:
    f.write(tmdl_content)

target_file = (
    script_dir
    / "SolArch.SemanticModel"
    / "definition"
    / "expressions.tmdl"
)

tmdl_content = r'''
expression Parameter1 = #"Sample File" meta [IsParameterQuery=true, BinaryIdentifier=#"Sample File", Type="Binary", IsParameterQueryRequired=true]
	lineageTag: 90e3bd6e-f852-4f27-9a3a-1b81938e80d2
	queryGroup: 'Transform File from Data\Helper Queries'

	annotation PBI_ResultType = Binary

expression 'Transform Sample File' =
		let
		    Source = Excel.Workbook(Parameter1, null, true),
		    Data_Sheet = Source{[Item="Data",Kind="Sheet"]}[Data],
		    #"Promoted Headers" = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true])
		in
		    #"Promoted Headers"
	lineageTag: 1e1daa66-544e-4e21-b78c-e44b37f1da4e
	queryGroup: 'Transform File from Data'

	annotation PBI_ResultType = Table

expression 'Sample File' =
		let
		    Source = Folder.Files("SOMEPATH"),
		    Navigation1 = Source{0}[Content]
		in
		    Navigation1
	lineageTag: 493a3884-e864-47d7-bc89-65c6bc9c683a
	queryGroup: 'Transform File from Data\Helper Queries'

	annotation PBI_ResultType = Binary

expression 'Transform File' =
		let
		    Source = (Parameter1 as binary) => let
		    Source = Excel.Workbook(Parameter1, null, true),
		    Data_Sheet = Source{[Item="Data",Kind="Sheet"]}[Data],
		    #"Promoted Headers" = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true])
		in
		    #"Promoted Headers"
		in
		    Source
	mAttributes: [ FunctionQueryBinding = "{""exemplarFormulaName"":""Transform Sample File""}" ]
	lineageTag: bfe562fa-bb63-4a1b-a5a2-2f04c6a4a5a0
	queryGroup: 'Transform File from Data\Helper Queries'

	annotation PBI_ResultType = Function

'''

tmdl_content = tmdl_content.replace(
    r'Source = Folder.Files("SOMEPATH"),',
    f'Source = Folder.Files("{pq_path}"),'
)

with open(target_file, "w", encoding="utf-8", newline="\n") as f:
    f.write(tmdl_content)
