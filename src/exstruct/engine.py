from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional, TextIO
from contextlib import contextmanager

from .core.integrate import extract_workbook
from .core import cells as _cells
from .core.cells import set_table_detection_params
from .io import (
    save_as_json,
    save_as_toon,
    save_as_yaml,
    save_print_area_views,
    save_sheets,
    serialize_workbook,
)
from .models import SheetData, WorkbookData
from .render import export_pdf, export_sheet_images

ExtractionMode = Literal["light", "standard", "verbose"]


@dataclass(frozen=True)
class StructOptions:
    """
    Extraction-time options for ExStructEngine.

    Attributes:
        mode: Extraction mode. One of "light", "standard", "verbose".
              - light: cells + table candidates only (no COM, shapes/charts empty)
              - standard: texted shapes + arrows + charts (if COM available)
              - verbose: all shapes (width/height), charts, table candidates
        table_params: Optional dict passed to `set_table_detection_params(**table_params)`
                      before extraction. Use this to tweak table detection heuristics
                      per engine instance without touching global state.
    """

    mode: ExtractionMode = "standard"
    table_params: Optional[dict] = None  # forwarded to set_table_detection_params if provided
    include_cell_links: Optional[bool] = None  # None -> auto: verbose=True, others=False


@dataclass(frozen=True)
class OutputOptions:
    """
    Output-time options for ExStructEngine.

    Attributes:
        fmt: Default export format. One of "json", "yaml", "yml", "toon".
        pretty: Whether to pretty-print JSON; default False (compact).
        indent: Explicit indent size. If None and pretty=True, indent=2 for JSON.
        include_rows: Include SheetData.rows in output (set False to drop).
        include_shapes: Include SheetData.shapes in output.
        include_charts: Include SheetData.charts in output.
        include_tables: Include SheetData.table_candidates in output.
        include_print_areas: Include SheetData.print_areas in output.
        include_shape_size: Include Shape.w/h in output (auto: verbose=True, others False).
        include_chart_size: Include Chart.w/h in output (auto: verbose=True, others False).
        sheets_dir: Optional directory to write per-sheet files (in the chosen fmt).
        print_areas_dir: Optional directory to write one file per print area (in the chosen fmt).
        stream: Optional default stream for stdout output when output_path is None.
    """

    fmt: Literal["json", "yaml", "yml", "toon"] = "json"
    pretty: bool = False
    indent: int | None = None
    include_rows: bool = True
    include_shapes: bool = True
    include_charts: bool = True
    include_tables: bool = True
    include_print_areas: bool | None = None  # None -> auto (light=False, others=True)
    include_shape_size: bool | None = None
    include_chart_size: bool | None = None
    sheets_dir: Path | None = None
    print_areas_dir: Path | None = None
    stream: TextIO | None = None


class ExStructEngine:
    """
    Configurable engine for ExStruct extraction and export.

    Instances are immutable; override options per call if needed.

    Key behaviors:
        - StructOptions: extraction mode and optional table detection params.
        - OutputOptions: serialization format/pretty-print, include/exclude filters, per-sheet/per-print-area output dirs, etc.
        - Main methods:
            extract(path, mode=None) -> WorkbookData
                - Modes: light/standard/verbose
                - light: COM-free; cells + tables + print areas only (shapes/charts empty)
            serialize(workbook, ...) -> str
                - Applies include_* filters, then serializes
            export(workbook, ...)
                - Writes to file/stdout; optionally per-sheet and per-print-area files
            process(file_path, ...)
                - One-shot extract→export (CLI equivalent), with optional PDF/PNG
    """

    def __init__(
        self,
        options: StructOptions | None = None,
        output: OutputOptions | None = None,
    ) -> None:
        self.options = options or StructOptions()
        self.output = output or OutputOptions()

    @staticmethod
    def from_defaults() -> "ExStructEngine":
        """Factory to create an engine with default options."""
        return ExStructEngine()

    def _apply_table_params(self) -> None:
        if self.options.table_params:
            set_table_detection_params(**self.options.table_params)

    @contextmanager
    def _table_params_scope(self):
        """
        Temporarily apply table_params and restore previous global config afterward.
        """
        if not self.options.table_params:
            yield
            return
        prev = dict(_cells._DETECTION_CONFIG)  # type: ignore[attr-defined]
        set_table_detection_params(**self.options.table_params)
        try:
            yield
        finally:
            set_table_detection_params(**prev)

    def _resolve_size_flags(self) -> tuple[bool, bool]:
        """
        Determine whether to include Shape/Chart size fields in output.
        Auto: verbose -> include, others -> exclude.
        """
        include_shape_size = (
            self.output.include_shape_size
            if self.output.include_shape_size is not None
            else self.options.mode == "verbose"
        )
        include_chart_size = (
            self.output.include_chart_size
            if self.output.include_chart_size is not None
            else self.options.mode == "verbose"
        )
        return include_shape_size, include_chart_size

    def _include_print_areas(self) -> bool:
        """
        Decide whether to include print areas in output.
        Auto: light -> False, others -> True.
        """
        if self.output.include_print_areas is None:
            return self.options.mode != "light"
        return self.output.include_print_areas

    def _filter_sheet(self, sheet: SheetData) -> SheetData:
        include_shape_size, include_chart_size = self._resolve_size_flags()
        include_print_areas = self._include_print_areas()
        return SheetData(
            rows=sheet.rows if self.output.include_rows else [],
            shapes=[
                s if include_shape_size else s.model_copy(update={"w": None, "h": None})
                for s in sheet.shapes
            ]
            if self.output.include_shapes
            else [],
            charts=[
                c if include_chart_size else c.model_copy(update={"w": None, "h": None})
                for c in sheet.charts
            ]
            if self.output.include_charts
            else [],
            table_candidates=sheet.table_candidates if self.output.include_tables else [],
            print_areas=sheet.print_areas if include_print_areas else [],
        )

    def _filter_workbook(self, wb: WorkbookData) -> WorkbookData:
        filtered = {name: self._filter_sheet(sheet) for name, sheet in wb.sheets.items()}
        return WorkbookData(book_name=wb.book_name, sheets=filtered)

    def extract(self, file_path: str | Path, *, mode: ExtractionMode | None = None) -> WorkbookData:
        """
        ワークブックを抽出して WorkbookData を返す。

        Args:
            file_path: .xlsx/.xlsm/.xls のパス
            mode: light/standard/verbose（未指定ならエンジンの StructOptions.mode）
                - light: COM なし。セル+テーブル+印刷範囲のみ。
                - standard: テキスト付き図形+矢印+チャート。印刷範囲あり。サイズは保持するがデフォルト出力では非表示。
                - verbose: 全図形（サイズ付き）+チャート（サイズ付き）。
        """
        chosen_mode = mode or self.options.mode
        if chosen_mode not in ("light", "standard", "verbose"):
            raise ValueError(f"Unsupported mode: {chosen_mode}")
        include_links = (
            self.options.include_cell_links
            if self.options.include_cell_links is not None
            else chosen_mode == "verbose"
        )
        include_print_areas = True  # lightでも印刷範囲は抽出する
        with self._table_params_scope():
            return extract_workbook(
                Path(file_path),
                mode=chosen_mode,
                include_cell_links=include_links,
                include_print_areas=include_print_areas,
            )

    def serialize(
        self,
        data: WorkbookData,
        *,
        fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
        pretty: Optional[bool] = None,
        indent: int | None = None,
    ) -> str:
        """
        WorkbookData を include/exclude フィルタ適用後に文字列化する。

        Args:
            fmt: json/yaml/yml/toon（未指定なら OutputOptions.fmt）
            pretty/indent: JSON 整形オプション
        """
        filtered = self._filter_workbook(data)
        use_fmt = (fmt or self.output.fmt)
        use_pretty = self.output.pretty if pretty is None else pretty
        use_indent = self.output.indent if indent is None else indent
        return serialize_workbook(filtered, fmt=use_fmt, pretty=use_pretty, indent=use_indent)

    def export(
        self,
        data: WorkbookData,
        output_path: Path | None = None,
        *,
        fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
        pretty: Optional[bool] = None,
        indent: int | None = None,
        sheets_dir: Path | None = None,
        print_areas_dir: Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        WorkbookData をファイルまたは標準出力に書き出す。

        - include_* フィルタ後のデータを使用
        - sheets_dir を指定するとシートごとの個別ファイルも出力
        - print_areas_dir を指定すると印刷範囲ごとの個別ファイルも出力（light モードではデフォルト無効）

        Args:
            output_path: None なら標準出力、それ以外はファイル書き込み
            fmt/pretty/indent: シリアライズ設定（未指定は OutputOptions から）
            sheets_dir: シートごとの出力先ディレクトリ
            print_areas_dir: 印刷範囲ごとの出力先ディレクトリ
            stream: output_path が None のときに上書きしたい IO
        """
        text = self.serialize(data, fmt=fmt, pretty=pretty, indent=indent)
        target_stream = stream or self.output.stream
        chosen_fmt = (fmt or self.output.fmt)
        chosen_sheets_dir = sheets_dir if sheets_dir is not None else self.output.sheets_dir
        chosen_print_areas_dir = (
            print_areas_dir
            if print_areas_dir is not None
            else self.output.print_areas_dir
        )

        if output_path is not None:
            output_path.write_text(text, encoding="utf-8")
        else:
            import sys

            stream_target = target_stream or sys.stdout
            stream_target.write(text)
            if not text.endswith("\n"):
                stream_target.write("\n")

        if chosen_sheets_dir is not None:
            filtered = self._filter_workbook(data)
            save_sheets(
                filtered,
                chosen_sheets_dir,
                fmt=chosen_fmt,
                pretty=self.output.pretty if pretty is None else pretty,
                indent=self.output.indent if indent is None else indent,
            )

        if chosen_print_areas_dir is not None:
            include_shape_size, include_chart_size = self._resolve_size_flags()
            if self._include_print_areas():
                filtered = self._filter_workbook(data)
                save_print_area_views(
                    filtered,
                    chosen_print_areas_dir,
                    fmt=chosen_fmt,
                    pretty=self.output.pretty if pretty is None else pretty,
                    indent=self.output.indent if indent is None else indent,
                    include_shapes=self.output.include_shapes,
                    include_charts=self.output.include_charts,
                    include_shape_size=include_shape_size,
                    include_chart_size=include_chart_size,
                )

        return None

    def process(
        self,
        file_path: Path,
        output_path: Path | None = None,
        *,
        out_fmt: Optional[str] = None,
        image: bool = False,
        pdf: bool = False,
        dpi: int = 72,
        mode: ExtractionMode | None = None,
        pretty: bool | None = None,
        indent: int | None = None,
        sheets_dir: Path | None = None,
        print_areas_dir: Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        抽出→出力の一括実行ラッパー（CLI 相当）。必要なら PDF/PNG も出力。

        Args:
            file_path: 入力 Excel
            output_path: None なら標準出力、それ以外はファイル
            out_fmt: json/yaml/yml/toon
            image/pdf: True で PNG/PDF を追加出力（Excel + pypdfium2 が必要）
            dpi: 画像出力時の DPI
            mode: 抽出モード（未指定ならエンジンの StructOptions.mode）
            pretty/indent: JSON 整形
            sheets_dir: シートごとの出力先
            print_areas_dir: 印刷範囲ごとの出力先
            stream: 標準出力時の IO を上書きしたい場合
        """
        wb = self.extract(file_path, mode=mode)
        chosen_fmt = out_fmt or self.output.fmt
        self.export(
            wb,
            output_path=output_path,
            fmt=chosen_fmt,  # type: ignore[arg-type]
            pretty=pretty,
            indent=indent,
            sheets_dir=sheets_dir,
            print_areas_dir=print_areas_dir,
            stream=stream,
        )

        if pdf or image:
            base_target = output_path or file_path.with_suffix(
                ".yaml" if chosen_fmt in ("yaml", "yml") else ".toon" if chosen_fmt == "toon" else ".json"
            )
            pdf_path = base_target.with_suffix(".pdf")
            export_pdf(file_path, pdf_path)
            if image:
                images_dir = pdf_path.parent / f"{pdf_path.stem}_images"
                export_sheet_images(file_path, images_dir, dpi=dpi)
