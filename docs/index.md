# ExStruct Documentation

![ExStruct Image](assets/icon.webp)

Welcome to the ExStruct docs. This site covers usage, CLI, and API reference for extracting structured data from Excel workbooks.

- [**README (EN)**](README.en.md)
- [**README (JA)**](README.ja.md)
- [**API Reference**](api.md): See the navigation for callable functions and parameters.
- [**CLI Guide**](cli.md)
- [**Concept / Why ExStruct?**](concept.md)
- [**Corporate Usage License Guide**](license-guide.md)

## At a Glance

- Export structured JSON (default), YAML, TOON (optional deps).
- Output modes: `light` / `standard` / `verbose`.
- Table detection is tunable via `set_table_detection_params`.
- Graceful fallback when Excel COM is unavailable.

## Quick Links

- Install: `pip install exstruct`
- CLI: `exstruct input.xlsx --pretty --mode standard`
- Python: `from exstruct import extract, export`
