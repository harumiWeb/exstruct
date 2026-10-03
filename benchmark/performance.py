"""Reproducible performance runner: uv run python -m benchmark.performance."""

from __future__ import annotations

import argparse
from contextlib import ExitStack, redirect_stdout
import ctypes
from datetime import UTC, datetime
import hashlib
from importlib.metadata import PackageNotFoundError, version
import json
import math
import os
from pathlib import Path
import platform
import statistics
import subprocess
import sys
from time import perf_counter
from typing import TYPE_CHECKING, Any, Literal

if TYPE_CHECKING:
    from exstruct.models import WorkbookData

HEAVY_MODULES = ("pandas", "numpy", "scipy", "xlwings", "openpyxl", "PIL", "pypdfium2")
BenchmarkMode = Literal["light", "standard"]
COLD_EXTRACT_SCRIPT = """
import sys
from contextlib import redirect_stdout
from time import perf_counter
import exstruct

started = perf_counter()
with redirect_stdout(sys.stderr):
    workbook = exstruct.extract(sys.argv[1], mode=sys.argv[2])
    exstruct.serialize_workbook(workbook, fmt="json")
print((perf_counter() - started) * 1000)
"""


def imported_modules() -> list[str]:
    return [name for name in HEAVY_MODULES if name in sys.modules]


def peak_rss_bytes() -> int | None:
    """Process lifetime RSS high-water mark; exclude Excel/other child processes."""
    if os.name == "nt":
        from ctypes import wintypes

        class MemoryCounters(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD),
                ("PageFaultCount", wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        psapi = ctypes.WinDLL("psapi", use_last_error=True)
        kernel.GetCurrentProcess.restype = wintypes.HANDLE
        psapi.GetProcessMemoryInfo.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(MemoryCounters),
            wintypes.DWORD,
        ]
        psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
        counters = MemoryCounters()
        counters.cb = ctypes.sizeof(counters)
        if psapi.GetProcessMemoryInfo(
            kernel.GetCurrentProcess(), ctypes.byref(counters), counters.cb
        ):
            return int(counters.PeakWorkingSetSize)
        return None
    try:
        import resource

        rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return int(rss if sys.platform == "darwin" else rss * 1024)
    except (ImportError, OSError):
        return None


def worker(
    path: Path, mode: BenchmarkMode, repeats: int, profile_all_features: bool = False
) -> dict[str, Any]:
    """Measure default public extraction before loading profiling utilities."""
    started = perf_counter()
    import exstruct

    import_ms = (perf_counter() - started) * 1000
    import_modules = imported_modules()
    samples = []
    with redirect_stdout(sys.stderr):
        for _ in range(repeats + 1):
            started = perf_counter()
            workbook = exstruct.extract(path, mode=mode)
            extract_ms = (perf_counter() - started) * 1000
            started = perf_counter()
            serialized = exstruct.serialize_workbook(workbook, fmt="json")
            serialize_ms = (perf_counter() - started) * 1000
            samples.append(
                {
                    "extract_ms": extract_ms,
                    "serialize_ms": serialize_ms,
                    "total_ms": extract_ms + serialize_ms,
                }
            )
        latency_peak = peak_rss_bytes()
        extraction_modules = imported_modules()

        from .performance_profile import StageRecorder

        recorder = StageRecorder()

        def profile_extract() -> WorkbookData:
            if profile_all_features:
                return exstruct.extract_workbook(
                    path,
                    mode=mode,
                    include_formulas_map=True,
                    include_colors_map=True,
                    include_merged_cells=True,
                )
            return exstruct.extract(path, mode=mode)

        expected = (
            exstruct.serialize_workbook(profile_extract(), fmt="json")
            if profile_all_features
            else serialized
        )
        with ExitStack() as stack:
            recorder.install(
                stack, direct_ooxml=mode == "light" and not profile_all_features
            )
            started = perf_counter()
            profiled_workbook = profile_extract()
            profile_ms = (perf_counter() - started) * 1000
        started = perf_counter()
        profile_serialized = exstruct.serialize_workbook(profiled_workbook, fmt="json")
        profile_serialize_ms = (perf_counter() - started) * 1000
        if expected != profile_serialized:
            raise RuntimeError(
                "Instrumented extraction differs from uninstrumented output"
            )
    return {
        "import_ms": import_ms,
        "modules_after_import": import_modules,
        "modules_after_extraction": extraction_modules,
        "first": samples[0],
        "repeated": samples[1:],
        "repeated_median": {
            key: statistics.median(sample[key] for sample in samples[1:])
            for key in samples[0]
        },
        "peak_rss_bytes": latency_peak,
        "output_bytes": len(serialized.encode("utf-8")),
        "profile": {
            "all_features": profile_all_features,
            "extract_ms": profile_ms,
            "serialize_ms": profile_serialize_ms,
            "total_ms": profile_ms + profile_serialize_ms,
            "stages": recorder.stages,
            "openpyxl_workbook_opens": recorder.stages["workbook_parsing"]["calls"],
            "archive_opens": recorder.archive_opens,
            "pipeline_state": recorder.state,
        },
    }


def fresh_process(arguments: list[str], timeout: float) -> tuple[float, str]:
    started = perf_counter()
    completed = subprocess.run(
        [sys.executable, *arguments],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=timeout,
        check=True,
    )
    elapsed_ms = (perf_counter() - started) * 1000
    if completed.stderr:
        print(completed.stderr, file=sys.stderr, end="")
    return elapsed_ms, completed.stdout


def source_metadata(timeout: float) -> dict[str, Any]:
    try:
        revision = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            check=True,
            capture_output=True,
            text=True,
            timeout=timeout,
        ).stdout.strip()
        dirty = bool(
            subprocess.run(
                ["git", "status", "--porcelain"],
                check=True,
                capture_output=True,
                text=True,
                timeout=timeout,
            ).stdout.strip()
        )
        return {"revision": revision, "dirty": dirty}
    except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return {"revision": None, "dirty": None}


def run_benchmark(
    path: Path,
    mode: BenchmarkMode,
    repeats: int,
    timeout: float,
    profile_all_features: bool = False,
) -> dict[str, Any]:
    path = path.resolve(strict=True)
    with path.open("rb") as stream:
        input_hash = hashlib.file_digest(stream, "sha256").hexdigest()
    startup_ms, _ = fresh_process(["-c", "pass"], timeout)
    bare_import_ms, _ = fresh_process(["-c", "import exstruct"], timeout)
    cli_ms, _ = fresh_process(["-m", "exstruct.cli.main", "--help"], timeout)
    worker_arguments = [
        "-m",
        "benchmark.performance",
        "--worker",
        "--input",
        str(path),
        "--mode",
        mode,
        "--repeats",
        str(repeats),
    ]
    if profile_all_features:
        worker_arguments.append("--profile-all-features")
    cold_ms, output = fresh_process(worker_arguments, timeout)
    measurements = json.loads(output)
    # Full child wall time includes all repeats and profile, so keep it distinct
    # from cold first extraction rather than falsely labelling it cold latency.
    cold_extract_ms, cold_output = fresh_process(
        [
            "-c",
            COLD_EXTRACT_SCRIPT,
            str(path),
            mode,
        ],
        timeout,
    )
    dependencies = {}
    for name in (
        "exstruct",
        "openpyxl",
        "pandas",
        "numpy",
        "scipy",
        "xlwings",
        "pydantic",
    ):
        try:
            dependencies[name] = version(name)
        except PackageNotFoundError:
            dependencies[name] = None
    return {
        "schema_version": 1,
        "recorded_at": datetime.now(UTC).isoformat(),
        "source": source_metadata(timeout),
        "environment": {
            "python": sys.version,
            "executable": sys.executable,
            "platform": platform.platform(),
            "dependencies": dependencies,
            "SKIP_COM_TESTS": os.environ.get("SKIP_COM_TESTS"),
            "EXSTRUCT_BORDER_CLUSTER_BACKEND": os.environ.get(
                "EXSTRUCT_BORDER_CLUSTER_BACKEND"
            ),
        },
        "input": {
            "name": path.name,
            "sha256": input_hash,
            "size_bytes": path.stat().st_size,
        },
        "mode": mode,
        "startup": {
            "python_process_ms": startup_ms,
            "package_import_process_ms": bare_import_ms,
            "cli_help_process_ms": cli_ms,
            "cold_extraction_process_ms": cold_extract_ms,
            "cold_extraction_in_process_ms": float(cold_output),
            "measurement_worker_process_ms": cold_ms,
        },
        **measurements,
    }


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return number


def positive_seconds(value: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise argparse.ArgumentTypeError("must be finite and greater than zero")
    return number


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--mode", choices=("light", "standard"), default="light")
    parser.add_argument("--repeats", type=positive_int, default=3)
    parser.add_argument("--timeout", type=positive_seconds, default=300.0)
    parser.add_argument("--output", type=Path)
    parser.add_argument(
        "--profile-all-features",
        action="store_true",
        help="Enable formulas, colors and merged cells only in the separate profile run",
    )
    parser.add_argument("--generate-fixtures", type=Path)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.generate_fixtures is not None:
        if args.input is not None or args.worker:
            parser.error("fixture generation must run separately from measurements")
        from .performance_fixtures import generate_fixtures

        generate_fixtures(args.generate_fixtures)
        return 0
    if args.input is None or not args.input.is_file():
        parser.error("--input must name an existing workbook")
    if args.output is not None and args.output.resolve() == args.input.resolve():
        parser.error("--output must differ from --input")
    try:
        result = (
            worker(args.input, args.mode, args.repeats, args.profile_all_features)
            if args.worker
            else run_benchmark(
                args.input,
                args.mode,
                args.repeats,
                args.timeout,
                args.profile_all_features,
            )
        )
        encoded = (
            json.dumps(result, indent=2, ensure_ascii=True, allow_nan=False) + "\n"
        )
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(encoded, encoding="utf-8")
        else:
            print(encoded, end="")
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"Benchmark failed: {exc}", file=sys.stderr)
        if isinstance(exc, subprocess.CalledProcessError):
            print(exc.stderr, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
