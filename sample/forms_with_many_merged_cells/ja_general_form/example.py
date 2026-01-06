from exstruct import ExStructEngine, StructOptions

file_path = "ja_form.xlsx"

engine = ExStructEngine(
    options=StructOptions(include_merged_values_in_rows=False),
)
wb = engine.extract(file_path, mode="standard")
engine.export(wb, "output.json", pretty=False)
