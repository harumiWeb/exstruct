from exstruct import (
    ColorsOptions,
    ExStructEngine,
    StructOptions,
)

file_path = "sample.xlsx"

engine = ExStructEngine(
    options=StructOptions(
        include_colors_map=True,
        colors=ColorsOptions(include_default_background=False),
    ),
)
wb = engine.extract(file_path)
engine.export(wb, "sample.json", pretty=True)
