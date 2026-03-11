"""LibreOffice runtime session management and UNO payload parsing helpers."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
import json
import math
import os
from pathlib import Path
import shutil
import socket
import subprocess  # nosec B404 - required for validated local runtime/process management
import sys
from tempfile import NamedTemporaryFile, mkdtemp
import time
from typing import Literal, TextIO, cast

_DEFAULT_STARTUP_TIMEOUT_SEC = 15.0
_DEFAULT_EXEC_TIMEOUT_SEC = 30.0
_DEFAULT_PYTHON_PROBE_TIMEOUT_SEC = 5.0
_STARTUP_BRIDGE_HANDSHAKE_TIMEOUT_SEC = 5.0
_STARTUP_PORT_RETRY_LIMIT = 3
_STARTUP_PORT_RETRY_BACKOFF_SEC = 0.1
_STDERR_SINK_UNLINK_TIMEOUT_SEC = 1.0
_STDERR_SINK_UNLINK_RETRY_INTERVAL_SEC = 0.1
_SUBPROCESS_ENV_ALLOWLIST = (
    "HOME",
    "LANG",
    "LC_ALL",
    "LC_CTYPE",
    "PATH",
    "SYSTEMROOT",
    "TEMP",
    "TMP",
    "USERPROFILE",
    "WINDIR",
)


class LibreOfficeUnavailableError(RuntimeError):
    """Raised when the LibreOffice runtime is not available."""


@dataclass(frozen=True)
class LibreOfficeChartGeometry:
    """Best-effort chart geometry captured from LibreOffice draw pages."""

    name: str
    persist_name: str | None = None
    left: int | None = None
    top: int | None = None
    width: int | None = None
    height: int | None = None


@dataclass(frozen=True)
class LibreOfficeDrawPageShape:
    """Best-effort shape snapshot captured from a LibreOffice draw page."""

    name: str
    shape_type: str | None = None
    text: str = ""
    left: int | None = None
    top: int | None = None
    width: int | None = None
    height: int | None = None
    rotation: float | None = None
    is_connector: bool = False
    start_shape_name: str | None = None
    end_shape_name: str | None = None


@dataclass(frozen=True)
class LibreOfficeSessionConfig:
    """Configuration required to launch and query a LibreOffice session."""

    soffice_path: Path
    startup_timeout_sec: float
    exec_timeout_sec: float
    profile_root: Path | None


@dataclass(frozen=True)
class _LibreOfficeStartupAttempt:
    """Configuration for one LibreOffice startup strategy."""

    name: str
    use_temp_profile: bool


@dataclass(frozen=True)
class _LibreOfficeStartupSuccess:
    """Process state returned by a successful LibreOffice startup attempt."""

    process: subprocess.Popen[str]
    port: int
    temp_profile_dir: Path | None
    stderr_sink: TextIO | None
    stderr_path: Path | None


@dataclass(frozen=True)
class _LibreOfficeStartupFailure:
    """User-facing detail captured for a failed LibreOffice startup attempt."""

    attempt_name: str
    message: str


class _LibreOfficeStartupAttemptError(RuntimeError):
    """Internal wrapper used to retry alternate LibreOffice startup strategies."""

    def __init__(self, failure: _LibreOfficeStartupFailure) -> None:
        """Store the structured startup failure for later aggregation."""

        super().__init__(failure.message)
        self.failure = failure


class LibreOfficeSession:
    """Best-effort runtime guard for LibreOffice-backed extraction."""

    def __init__(self, config: LibreOfficeSessionConfig) -> None:
        """Initialize a session wrapper for a specific LibreOffice runtime configuration."""

        self.config = config
        self._temp_profile_dir: Path | None = None
        self._soffice_process: subprocess.Popen[str] | None = None
        self._accept_port: int | None = None
        self._python_path = _resolve_python_path(config.soffice_path)
        self._bridge_payload_cache: dict[str, object] = {}
        self._soffice_stderr_sink: TextIO | None = None
        self._soffice_stderr_path: Path | None = None

    @classmethod
    def from_env(cls) -> LibreOfficeSession:
        """Build a session from ExStruct environment variables."""
        raw_path = os.getenv("EXSTRUCT_LIBREOFFICE_PATH")
        resolved = (
            _validated_runtime_path(Path(raw_path)) if raw_path else _which_soffice()
        )
        if resolved is None:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: soffice was not found."
            )
        return cls(
            LibreOfficeSessionConfig(
                soffice_path=resolved,
                startup_timeout_sec=_get_timeout_from_env(
                    "EXSTRUCT_LIBREOFFICE_STARTUP_TIMEOUT_SEC",
                    default=_DEFAULT_STARTUP_TIMEOUT_SEC,
                ),
                exec_timeout_sec=_get_timeout_from_env(
                    "EXSTRUCT_LIBREOFFICE_EXEC_TIMEOUT_SEC",
                    default=_DEFAULT_EXEC_TIMEOUT_SEC,
                ),
                profile_root=_get_optional_path("EXSTRUCT_LIBREOFFICE_PROFILE_ROOT"),
            )
        )

    def __enter__(self) -> LibreOfficeSession:
        """Launch a headless LibreOffice process and wait for its UNO socket.

        Returns:
            The initialized session instance.

        Raises:
            LibreOfficeUnavailableError: If the runtime executable, compatible Python runtime, version probe, or UNO socket startup is unavailable.
        """

        if not self.config.soffice_path.exists():
            raise LibreOfficeUnavailableError(
                f"LibreOffice runtime is unavailable: '{self.config.soffice_path}' was not found."
            )
        if self._python_path is None or not self._python_path.exists():
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: compatible Python runtime was not found."
            )
        _probe_soffice_runtime(
            soffice_path=self.config.soffice_path,
            timeout_sec=self.config.startup_timeout_sec,
        )
        failures: list[_LibreOfficeStartupFailure] = []
        for attempt in _iter_startup_attempts():
            try:
                startup = _start_soffice_startup_attempt(
                    soffice_path=self.config.soffice_path,
                    python_path=self._python_path,
                    profile_root=self.config.profile_root,
                    startup_timeout_sec=self.config.startup_timeout_sec,
                    attempt=attempt,
                )
            except _LibreOfficeStartupAttemptError as exc:
                failures.append(exc.failure)
                continue
            self._soffice_process = startup.process
            self._accept_port = startup.port
            self._temp_profile_dir = startup.temp_profile_dir
            self._soffice_stderr_sink = startup.stderr_sink
            self._soffice_stderr_path = startup.stderr_path
            return self
        raise LibreOfficeUnavailableError(_format_startup_failures(failures))

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        """Terminate the LibreOffice process and remove the temporary user profile."""

        _ = exc_type
        _ = exc
        _ = tb
        if self._soffice_process is not None:
            _shutdown_soffice_process(self._soffice_process)
            self._soffice_process = None
        _close_stderr_sink(self._soffice_stderr_sink, self._soffice_stderr_path)
        self._soffice_stderr_sink = None
        self._soffice_stderr_path = None
        self._accept_port = None
        if self._temp_profile_dir is not None:
            _cleanup_profile_dir(self._temp_profile_dir)
            self._temp_profile_dir = None

    def load_workbook(self, file_path: Path) -> object:
        """Return a lightweight workbook token for future subprocess integration."""
        return {"file_path": str(file_path.resolve())}

    def close_workbook(self, workbook: object) -> None:
        """Close a workbook token returned by ``load_workbook``."""
        _ = workbook

    def extract_draw_page_shapes(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Extract best-effort draw-page shapes from the workbook."""
        payload = self._run_bridge(file_path, kind="draw-page")
        return _parse_draw_page_payload(payload)

    def extract_chart_geometries(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Extract chart geometry candidates from the workbook draw pages."""
        payload = self._run_bridge(file_path, kind="charts")
        return _parse_chart_payload(payload)

    def _run_bridge(
        self,
        file_path: Path,
        *,
        kind: Literal["charts", "draw-page"],
    ) -> object:
        """Run the bundled bridge script and cache the parsed payload.

        Args:
            file_path: Workbook to inspect through LibreOffice.
            kind: Extraction kind to request from the bridge script.

        Returns:
            Decoded JSON payload returned by the bridge script.

        Raises:
            LibreOfficeUnavailableError: If the compatible Python runtime or bridge invocation is unavailable or times out.
            RuntimeError: If the bridge fails or returns invalid JSON.
        """

        if self._accept_port is None or self._python_path is None:
            raise RuntimeError("LibreOfficeSession must be entered before extraction.")
        cache_key = f"{kind}:{file_path.resolve()}"
        if cache_key in self._bridge_payload_cache:
            return self._bridge_payload_cache[cache_key]
        try:
            completed = _run_bridge_extract_subprocess(
                python_path=self._python_path,
                host="127.0.0.1",
                port=self._accept_port,
                file_path=file_path.resolve(),
                kind=kind,
                timeout_sec=self.config.exec_timeout_sec,
            )
        except FileNotFoundError as exc:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: compatible Python runtime could not be executed."
            ) from exc
        except subprocess.TimeoutExpired as exc:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: UNO bridge extraction timed out."
            ) from exc
        except subprocess.CalledProcessError as exc:
            detail = exc.stderr.strip() or exc.stdout.strip() or "unknown error"
            raise RuntimeError(
                f"LibreOffice UNO bridge extraction failed: {detail}"
            ) from exc
        try:
            payload = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "LibreOffice UNO bridge extraction returned invalid JSON."
            ) from exc
        self._bridge_payload_cache[cache_key] = payload
        return payload


def _which_soffice() -> Path | None:
    """Return the first discoverable ``soffice`` executable on ``PATH``."""

    for candidate in ("soffice", "soffice.com", "soffice.exe"):
        resolved = shutil.which(candidate)
        if resolved:
            return _validated_runtime_path(Path(resolved))
    return None


def _get_timeout_from_env(name: str, *, default: float) -> float:
    """Read a positive finite timeout value from the environment."""

    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        value = float(raw)
    except ValueError:
        raise ValueError(f"{name} must be a positive finite float.") from None
    if not math.isfinite(value) or value <= 0:
        raise ValueError(f"{name} must be a positive finite float.")
    return value


def _get_optional_path(name: str) -> Path | None:
    """Return an optional path from the environment when set."""

    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return None
    return Path(raw)


def _probe_soffice_runtime(*, soffice_path: Path, timeout_sec: float) -> None:
    """Verify that the configured soffice executable is runnable."""

    try:
        _run_soffice_version_subprocess(
            soffice_path=soffice_path, timeout_sec=timeout_sec
        )
    except FileNotFoundError as exc:
        raise LibreOfficeUnavailableError(
            f"LibreOffice runtime is unavailable: '{soffice_path}' could not be executed."
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: soffice version probe timed out."
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: soffice version probe failed."
        ) from exc


def _iter_startup_attempts() -> tuple[_LibreOfficeStartupAttempt, ...]:
    """Return the ordered startup strategies used for the LibreOffice session."""

    return (
        _LibreOfficeStartupAttempt(name="isolated-profile", use_temp_profile=True),
        _LibreOfficeStartupAttempt(name="shared-profile", use_temp_profile=False),
    )


def _start_soffice_startup_attempt(
    *,
    soffice_path: Path,
    python_path: Path,
    profile_root: Path | None,
    startup_timeout_sec: float,
    attempt: _LibreOfficeStartupAttempt,
) -> _LibreOfficeStartupSuccess:
    """Launch soffice for one startup strategy and wait for the UNO socket."""

    temp_profile_dir = (
        _create_temp_profile_dir(profile_root) if attempt.use_temp_profile else None
    )
    failure_messages: list[str] = []
    for startup_index in range(1, _STARTUP_PORT_RETRY_LIMIT + 1):
        port = _reserve_tcp_port()
        process: subprocess.Popen[str] | None = None
        stderr_sink: TextIO | None = None
        stderr_path: Path | None = None
        try:
            stderr_sink, stderr_path = _create_stderr_sink()
            process = _spawn_trusted_subprocess(
                _build_soffice_startup_command(
                    soffice_path=soffice_path,
                    port=port,
                    temp_profile_dir=temp_profile_dir,
                ),
                stdout=subprocess.DEVNULL,
                stderr=stderr_sink,
            )
            _wait_for_socket(
                host="127.0.0.1",
                port=port,
                timeout_sec=startup_timeout_sec,
                process=process,
            )
            _probe_uno_bridge_handshake(
                python_path=python_path,
                host="127.0.0.1",
                port=port,
                timeout_sec=min(
                    startup_timeout_sec,
                    _STARTUP_BRIDGE_HANDSHAKE_TIMEOUT_SEC,
                ),
                process=process,
            )
        except FileNotFoundError as exc:
            detail = _cleanup_failed_startup_process(
                process=process,
                stderr_sink=stderr_sink,
                stderr_path=stderr_path,
            )
            if temp_profile_dir is not None:
                _cleanup_profile_dir(temp_profile_dir)
            raise _LibreOfficeStartupAttemptError(
                _LibreOfficeStartupFailure(
                    attempt_name=attempt.name,
                    message=_append_startup_detail(
                        f"'{soffice_path}' could not be executed.",
                        detail,
                    ),
                )
            ) from exc
        except OSError as exc:
            detail = _cleanup_failed_startup_process(
                process=process,
                stderr_sink=stderr_sink,
                stderr_path=stderr_path,
            )
            if temp_profile_dir is not None:
                _cleanup_profile_dir(temp_profile_dir)
            raise _LibreOfficeStartupAttemptError(
                _LibreOfficeStartupFailure(
                    attempt_name=attempt.name,
                    message=_append_startup_detail(
                        f"soffice startup could not be launched ({exc.__class__.__name__}: {exc}).",
                        detail,
                    ),
                )
            ) from exc
        except LibreOfficeUnavailableError as exc:
            detail = _cleanup_failed_startup_process(
                process=process,
                stderr_sink=stderr_sink,
                stderr_path=stderr_path,
            )
            failure_messages.append(
                _append_startup_detail(
                    _strip_runtime_unavailable_prefix(str(exc)),
                    detail,
                )
            )
            if (
                startup_index >= _STARTUP_PORT_RETRY_LIMIT
                or not _should_retry_startup_failure(failure_messages[-1])
            ):
                if temp_profile_dir is not None:
                    _cleanup_profile_dir(temp_profile_dir)
                raise _LibreOfficeStartupAttemptError(
                    _LibreOfficeStartupFailure(
                        attempt_name=attempt.name,
                        message=_format_startup_retry_failures(failure_messages),
                    )
                ) from exc
            time.sleep(_STARTUP_PORT_RETRY_BACKOFF_SEC)
            continue
        if process is None:
            raise RuntimeError("LibreOffice startup attempt did not create a process.")
        return _LibreOfficeStartupSuccess(
            process=process,
            port=port,
            temp_profile_dir=temp_profile_dir,
            stderr_sink=stderr_sink,
            stderr_path=stderr_path,
        )
    if temp_profile_dir is not None:
        _cleanup_profile_dir(temp_profile_dir)
    raise _LibreOfficeStartupAttemptError(
        _LibreOfficeStartupFailure(
            attempt_name=attempt.name,
            message=_format_startup_retry_failures(failure_messages),
        )
    )


def _create_temp_profile_dir(profile_root: Path | None) -> Path:
    """Create a temporary LibreOffice profile directory for an isolated launch."""

    if profile_root is not None:
        profile_root.mkdir(parents=True, exist_ok=True)
        return Path(mkdtemp(prefix="exstruct-lo-", dir=str(profile_root)))
    return Path(mkdtemp(prefix="exstruct-lo-"))


def _build_soffice_startup_command(
    *,
    soffice_path: Path,
    port: int,
    temp_profile_dir: Path | None,
) -> list[str]:
    """Build the soffice command used for a startup attempt."""

    args = [
        _subprocess_executable_arg(soffice_path),
        "--headless",
        "--nologo",
        "--nodefault",
        "--norestore",
        "--nolockcheck",
    ]
    if temp_profile_dir is not None:
        args.append(f"-env:UserInstallation={temp_profile_dir.as_uri()}")
    args.append(
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    )
    return args


def _cleanup_failed_startup_process(
    *,
    process: subprocess.Popen[str] | None,
    stderr_sink: TextIO | None,
    stderr_path: Path | None,
) -> str | None:
    """Terminate a failed startup process and return a bounded stderr detail."""

    detail: str | None = None
    if process is not None:
        _shutdown_soffice_process(process)
        detail = _read_soffice_startup_stderr(
            process=process,
            stderr_sink=stderr_sink,
        )
    _close_stderr_sink(stderr_sink, stderr_path)
    return detail


def _shutdown_soffice_process(process: subprocess.Popen[str]) -> None:
    """Terminate a soffice process with a bounded force-kill fallback."""

    if process.poll() is not None:
        try:
            process.wait(timeout=0.2)
        except subprocess.TimeoutExpired:
            pass
        return
    process.terminate()
    try:
        process.wait(timeout=5.0)
    except subprocess.TimeoutExpired:
        process.kill()
        try:
            process.wait(timeout=5.0)
        except subprocess.TimeoutExpired:
            pass


def _create_stderr_sink() -> tuple[TextIO, Path]:
    """Create a temporary sink file for long-running soffice stderr."""

    sink = NamedTemporaryFile(
        mode="w+",
        encoding="utf-8",
        prefix="exstruct-soffice-",
        suffix=".log",
        delete=False,
    )
    return (cast(TextIO, sink), Path(sink.name))


def _close_stderr_sink(stderr_sink: TextIO | None, stderr_path: Path | None) -> None:
    """Close and remove a temporary soffice stderr sink."""

    if stderr_sink is not None:
        stderr_sink.close()
    if stderr_path is not None:
        _unlink_stderr_sink_path(stderr_path)


def _unlink_stderr_sink_path(path: Path) -> None:
    """Best-effort unlink for temporary stderr logs without masking prior failures."""

    deadline = time.monotonic() + _STDERR_SINK_UNLINK_TIMEOUT_SEC
    while True:
        try:
            path.unlink()
            return
        except FileNotFoundError:
            return
        except PermissionError:
            if time.monotonic() >= deadline:
                return
            time.sleep(_STDERR_SINK_UNLINK_RETRY_INTERVAL_SEC)


def _read_soffice_startup_stderr(
    *,
    process: subprocess.Popen[str],
    stderr_sink: TextIO | None,
) -> str | None:
    """Read a short stderr snippet from a failed soffice startup attempt."""

    stderr = ""
    if stderr_sink is not None:
        stderr_sink.flush()
        stderr_sink.seek(0)
        stderr = stderr_sink.read()
    if not stderr:
        try:
            _stdout, stderr = process.communicate(timeout=0.2)
        except subprocess.TimeoutExpired:
            return None
    if not stderr:
        return None
    cleaned = " ".join(str(stderr).strip().split())
    if not cleaned:
        return None
    return f"stderr={cleaned[:240]}"


def _append_startup_detail(message: str, detail: str | None) -> str:
    """Append bounded subprocess detail to a startup message when available."""

    if detail is None:
        return message
    return f"{message} {detail}"


def _should_retry_startup_failure(message: str) -> bool:
    """Return True when a startup failure looks consistent with port contention."""

    lowered = message.lower()
    retryable_markers = (
        "soffice socket startup timed out.",
        "soffice exited during startup.",
        "uno bridge handshake timed out.",
        "uno bridge handshake failed.",
        "address already in use",
        "already in use",
        "failed to bind",
        "cannot bind",
        "could not bind",
        "listen",
    )
    return any(marker in lowered for marker in retryable_markers)


def _format_startup_retry_failures(failure_messages: Sequence[str]) -> str:
    """Render numbered startup failures captured within one strategy."""

    if not failure_messages:
        return "soffice startup failed."
    return "; ".join(
        f"attempt {index}/{_STARTUP_PORT_RETRY_LIMIT}: {message}"
        for index, message in enumerate(failure_messages, start=1)
    )


def _strip_runtime_unavailable_prefix(message: str) -> str:
    """Strip the shared public error prefix from an internal startup message."""

    prefix = "LibreOffice runtime is unavailable: "
    if message.startswith(prefix):
        return message[len(prefix) :]
    return message


def _format_startup_failures(
    failures: list[_LibreOfficeStartupFailure],
) -> str:
    """Render a combined startup failure message across all launch strategies."""

    if not failures:
        return "LibreOffice runtime is unavailable: soffice startup failed."
    detail = "; ".join(
        f"{failure.attempt_name}: {failure.message}" for failure in failures
    )
    return f"LibreOffice runtime is unavailable: soffice startup failed. ({detail})"


def _resolve_python_path(soffice_path: Path) -> Path | None:
    """Resolve a Python executable capable of running the LibreOffice bridge."""

    override = os.getenv("EXSTRUCT_LIBREOFFICE_PYTHON_PATH")
    if override:
        override_path = _validated_runtime_path(Path(override))
        detail = _probe_libreoffice_bridge_failure(override_path)
        if detail is not None:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: configured "
                "EXSTRUCT_LIBREOFFICE_PYTHON_PATH is incompatible with the "
                f"bundled bridge. ({detail})"
            )
        return override_path
    for program_dir in _soffice_program_dirs(soffice_path):
        for path in _bundled_python_candidates(program_dir):
            if _python_supports_libreoffice_bridge(path):
                return path
    for python_candidate in _system_python_candidates():
        if _python_supports_libreoffice_bridge(python_candidate):
            return python_candidate
    return None


def _soffice_program_dirs(soffice_path: Path) -> tuple[Path, ...]:
    """Return candidate LibreOffice program directories for the given ``soffice`` path."""

    program_dirs = [soffice_path.parent]
    try:
        resolved_parent = soffice_path.resolve(strict=False).parent
    except OSError:
        return tuple(program_dirs)
    if resolved_parent not in program_dirs:
        program_dirs.append(resolved_parent)
    return tuple(program_dirs)


def _bundled_python_candidates(program_dir: Path) -> tuple[Path, ...]:
    """Return bundled LibreOffice Python candidates for a program directory."""

    candidates: list[Path] = []
    seen_paths: set[Path] = set()
    for file_name in ("python.exe", "python.bin", "python"):
        path = program_dir / file_name
        if path.exists() and path not in seen_paths:
            candidates.append(path)
            seen_paths.add(path)
    try:
        child_dirs = tuple(
            child
            for child in program_dir.iterdir()
            if child.is_dir() and child.name.startswith("python-core-")
        )
    except OSError:
        return tuple(candidates)
    for child_dir in child_dirs:
        for relative_path in (
            Path("python.exe"),
            Path("python"),
            Path("bin/python.exe"),
            Path("bin/python"),
        ):
            bundled_path = child_dir / relative_path
            if bundled_path.exists() and bundled_path not in seen_paths:
                candidates.append(bundled_path)
                seen_paths.add(bundled_path)
    return tuple(candidates)


def _system_python_candidates() -> tuple[Path, ...]:
    """Return candidate system Python executables for Debian/Ubuntu-style setups."""

    candidates: list[Path] = []
    for raw_candidate in (
        sys.executable,
        shutil.which("python3"),
        shutil.which("python"),
    ):
        if not raw_candidate:
            continue
        path = Path(raw_candidate)
        if path not in candidates:
            candidates.append(path)
    for path in (Path("/usr/bin/python3"), Path("/usr/bin/python")):
        if path not in candidates:
            candidates.append(path)
    return tuple(candidates)


def _python_supports_libreoffice_bridge(python_path: Path) -> bool:
    """Return True when the candidate Python can run the bundled bridge probe."""

    return _probe_libreoffice_bridge_failure(python_path) is None


def _probe_libreoffice_bridge_failure(python_path: Path) -> str | None:
    """Return ``None`` on success, otherwise a short incompatibility detail."""

    if not python_path.exists():
        return f"'{python_path}' was not found."
    try:
        _run_bridge_probe_subprocess(
            python_path=python_path,
            timeout_sec=_DEFAULT_PYTHON_PROBE_TIMEOUT_SEC,
        )
    except OSError:
        return f"'{python_path}' could not be executed."
    except subprocess.TimeoutExpired:
        return "bundled bridge probe timed out."
    except subprocess.CalledProcessError as exc:
        return (
            exc.stderr.strip() or exc.stdout.strip() or "bundled bridge probe failed."
        )
    return None


def _reserve_tcp_port() -> int:
    """Return an ephemeral localhost TCP port candidate for the UNO socket."""

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _validated_runtime_path(path: Path) -> Path:
    """Return a normalized runtime path before it is used in subprocess argv."""

    normalized_path = _prefer_windows_console_soffice(path)
    try:
        return normalized_path.resolve(strict=False)
    except OSError:
        return normalized_path


def _prefer_windows_console_soffice(path: Path) -> Path:
    """Prefer ``soffice.com`` when a Windows caller points at ``soffice.exe``."""

    if sys.platform != "win32" or path.name.lower() != "soffice.exe":
        return path
    console_launcher = path.with_name("soffice.com")
    if console_launcher.exists():
        return console_launcher
    return path


def _subprocess_executable_arg(path: Path) -> str:
    """Return a validated executable argument for subprocess calls."""

    return str(_validated_runtime_path(path))


def _subprocess_path_arg(path: Path) -> str:
    """Return a normalized non-executable path argument for subprocess calls."""

    return str(path.resolve(strict=False))


def _bridge_subprocess_cwd(python_path: Path) -> Path:
    """Return the working directory used for LibreOffice bridge subprocesses."""

    return _validated_runtime_path(python_path).parent


def _build_subprocess_env(
    *,
    pythonioencoding: str | None = None,
    runtime_dirs: Sequence[Path] = (),
) -> dict[str, str]:
    """Return a minimal inherited environment for bridge-related subprocesses."""

    env: dict[str, str] = {}
    for key in _SUBPROCESS_ENV_ALLOWLIST:
        value = os.environ.get(key)
        if value:
            env[key] = value
    runtime_paths = [_subprocess_path_arg(path) for path in runtime_dirs]
    if runtime_paths:
        inherited_path = env.get("PATH")
        env["PATH"] = os.pathsep.join(
            [*runtime_paths, *([inherited_path] if inherited_path else [])]
        )
    if pythonioencoding is not None:
        env["PYTHONIOENCODING"] = pythonioencoding
    return env


def _run_soffice_version_subprocess(
    *,
    soffice_path: Path,
    timeout_sec: float,
) -> subprocess.CompletedProcess[str]:
    """Run `soffice --version` with a fixed argv shape."""

    # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
    # Safe by construction: executable path is validated locally and invoked via
    # shell=False with no user-controlled command string assembly.
    return subprocess.run(  # nosec B603  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
        [_subprocess_executable_arg(soffice_path), "--version"],
        capture_output=True,
        check=True,
        text=True,
        encoding="utf-8",
        timeout=timeout_sec,
    )


def _run_bridge_probe_subprocess(
    *,
    python_path: Path,
    timeout_sec: float,
) -> subprocess.CompletedProcess[str]:
    """Run the bundled LibreOffice bridge in `--probe` mode."""

    bridge_path = Path(__file__).with_name("_libreoffice_bridge.py")
    runtime_dir = _bridge_subprocess_cwd(python_path)
    env = _build_subprocess_env(pythonioencoding="utf-8", runtime_dirs=(runtime_dir,))
    # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
    # Safe by construction: validated Python executable + bundled local script +
    # fixed probe flag under shell=False. Probe mode forwards only the shared
    # allowlisted environment, while prepending the runtime dir needed for
    # Windows-hosted UNO imports and DLL resolution.
    return subprocess.run(  # nosec B603  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
        [
            _subprocess_executable_arg(python_path),
            _subprocess_path_arg(bridge_path),
            "--probe",
        ],
        capture_output=True,
        check=True,
        text=True,
        encoding="utf-8",
        timeout=timeout_sec,
        env=env,
        cwd=_bridge_subprocess_cwd(python_path),
    )


def _run_bridge_extract_subprocess(
    *,
    python_path: Path,
    host: str,
    port: int,
    file_path: Path,
    kind: Literal["charts", "draw-page"],
    timeout_sec: float,
) -> subprocess.CompletedProcess[str]:
    """Run the bundled LibreOffice bridge for workbook extraction."""

    bridge_path = Path(__file__).with_name("_libreoffice_bridge.py")
    runtime_dir = _bridge_subprocess_cwd(python_path)
    env = _build_subprocess_env(pythonioencoding="utf-8", runtime_dirs=(runtime_dir,))
    input_text = _subprocess_path_arg(file_path)
    # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
    # Safe by construction: executable/script paths are local validated files, the
    # workbook path is forwarded via stdin rather than command text, and no shell
    # command string is constructed. Host/port remain discrete argv entries.
    return subprocess.run(  # nosec B603  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
        [
            _subprocess_executable_arg(python_path),
            _subprocess_path_arg(bridge_path),
            "--host",
            host,
            "--port",
            str(port),
            "--file-stdin",
            "--kind",
            kind,
        ],
        capture_output=True,
        check=True,
        input=input_text,
        text=True,
        encoding="utf-8",
        timeout=timeout_sec,
        env=env,
        cwd=runtime_dir,
    )


def _run_bridge_handshake_subprocess(
    *,
    python_path: Path,
    host: str,
    port: int,
    timeout_sec: float,
) -> subprocess.CompletedProcess[str]:
    """Run the bundled LibreOffice bridge in handshake mode."""

    bridge_path = Path(__file__).with_name("_libreoffice_bridge.py")
    runtime_dir = _bridge_subprocess_cwd(python_path)
    env = _build_subprocess_env(pythonioencoding="utf-8", runtime_dirs=(runtime_dir,))
    # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
    # Safe by construction: validated Python executable + bundled local script +
    # fixed handshake flags under shell=False. Host/port remain argv elements, not
    # shell-expanded command text.
    return subprocess.run(  # nosec B603  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
        [
            _subprocess_executable_arg(python_path),
            _subprocess_path_arg(bridge_path),
            "--handshake",
            "--host",
            host,
            "--port",
            str(port),
            "--connect-timeout",
            str(timeout_sec),
        ],
        capture_output=True,
        check=True,
        text=True,
        encoding="utf-8",
        timeout=timeout_sec,
        env=env,
        cwd=runtime_dir,
    )


def _spawn_trusted_subprocess(
    args: Sequence[str],
    *,
    stdout: int | TextIO,
    stderr: int | TextIO,
) -> subprocess.Popen[str]:
    """Spawn a trusted long-running subprocess with fixed argv structure."""

    # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
    # Safe by construction: argv is a fixed `soffice` command assembled from validated
    # local paths and an ephemeral localhost port; no shell expansion is used.
    return subprocess.Popen(  # nosec B603  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit
        list(args),
        stdout=stdout,
        stderr=stderr,
        text=True,
    )


def _wait_for_socket(
    *,
    host: str,
    port: int,
    timeout_sec: float,
    process: subprocess.Popen[str] | None,
) -> None:
    """Wait for the LibreOffice UNO socket to accept connections."""

    deadline = time.monotonic() + timeout_sec
    while time.monotonic() < deadline:
        if process is not None and process.poll() is not None:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: soffice exited during startup."
            )
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            try:
                sock.connect((host, port))
            except OSError:
                time.sleep(0.1)
                continue
        return
    raise LibreOfficeUnavailableError(
        "LibreOffice runtime is unavailable: soffice socket startup timed out."
    )


def _probe_uno_bridge_handshake(
    *,
    python_path: Path,
    host: str,
    port: int,
    timeout_sec: float,
    process: subprocess.Popen[str] | None,
) -> None:
    """Verify that the accepting socket resolves a LibreOffice UNO context."""

    if process is not None and process.poll() is not None:
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: soffice exited during startup."
        )
    try:
        _run_bridge_handshake_subprocess(
            python_path=python_path,
            host=host,
            port=port,
            timeout_sec=timeout_sec,
        )
    except FileNotFoundError as exc:
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: compatible Python runtime could not be executed."
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: UNO bridge handshake timed out."
        ) from exc
    except subprocess.CalledProcessError as exc:
        detail = exc.stderr.strip() or exc.stdout.strip() or "unknown error"
        raise LibreOfficeUnavailableError(
            f"LibreOffice runtime is unavailable: UNO bridge handshake failed. ({detail})"
        ) from exc


def _cleanup_profile_dir(path: Path) -> None:
    """Retry cleanup of a temporary LibreOffice profile directory."""

    deadline = time.monotonic() + 5.0
    while True:
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists() or time.monotonic() >= deadline:
            return
        time.sleep(0.1)


def _parse_chart_payload(
    payload: object,
) -> dict[str, list[LibreOfficeChartGeometry]]:
    """Validate and coerce a raw bridge payload into chart geometries by sheet."""

    if not isinstance(payload, dict):
        raise RuntimeError(
            "LibreOffice UNO chart extraction returned a non-object payload."
        )
    sheets_payload = payload.get("sheets")
    if not isinstance(sheets_payload, dict):
        raise RuntimeError(
            "LibreOffice UNO chart extraction payload is missing sheets."
        )
    result: dict[str, list[LibreOfficeChartGeometry]] = {}
    for sheet_name, items in sheets_payload.items():
        if not isinstance(sheet_name, str) or not isinstance(items, list):
            continue
        geometries: list[LibreOfficeChartGeometry] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            name = item.get("name")
            if not isinstance(name, str) or not name:
                continue
            persist_name = item.get("persist_name")
            geometries.append(
                LibreOfficeChartGeometry(
                    name=name,
                    persist_name=persist_name
                    if isinstance(persist_name, str)
                    else None,
                    left=_coerce_int(item.get("left")),
                    top=_coerce_int(item.get("top")),
                    width=_coerce_int(item.get("width")),
                    height=_coerce_int(item.get("height")),
                )
            )
        result[sheet_name] = geometries
    return result


def _parse_draw_page_payload(
    payload: object,
) -> dict[str, list[LibreOfficeDrawPageShape]]:
    """Validate and coerce a raw bridge payload into draw-page shapes by sheet."""

    if not isinstance(payload, dict):
        raise RuntimeError(
            "LibreOffice UNO draw-page extraction returned a non-object payload."
        )
    sheets_payload = payload.get("sheets")
    if not isinstance(sheets_payload, dict):
        raise RuntimeError(
            "LibreOffice UNO draw-page extraction payload is missing sheets."
        )
    result: dict[str, list[LibreOfficeDrawPageShape]] = {}
    for sheet_name, items in sheets_payload.items():
        if not isinstance(sheet_name, str) or not isinstance(items, list):
            continue
        shapes: list[LibreOfficeDrawPageShape] = []
        for index, item in enumerate(items, start=1):
            if not isinstance(item, dict):
                continue
            raw_name = item.get("name")
            name = (
                raw_name if isinstance(raw_name, str) and raw_name else f"Shape {index}"
            )
            shape_type = item.get("shape_type")
            start_shape_name = item.get("start_shape_name")
            end_shape_name = item.get("end_shape_name")
            text = item.get("text")
            shapes.append(
                LibreOfficeDrawPageShape(
                    name=name,
                    shape_type=shape_type if isinstance(shape_type, str) else None,
                    text=text if isinstance(text, str) else "",
                    left=_coerce_int(item.get("left")),
                    top=_coerce_int(item.get("top")),
                    width=_coerce_int(item.get("width")),
                    height=_coerce_int(item.get("height")),
                    rotation=_coerce_rotation(item.get("rotation")),
                    is_connector=bool(item.get("is_connector")),
                    start_shape_name=start_shape_name
                    if isinstance(start_shape_name, str)
                    else None,
                    end_shape_name=end_shape_name
                    if isinstance(end_shape_name, str)
                    else None,
                )
            )
        result[sheet_name] = shapes
    return result


def _coerce_int(value: object) -> int | None:
    """Coerce numeric payload values to integers."""

    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value))
    return None


def _coerce_rotation(value: object) -> float | None:
    """Coerce numeric payload values to floating-point rotation angles."""

    if isinstance(value, int | float):
        return float(value)
    return None
