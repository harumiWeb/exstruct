from __future__ import annotations

from dataclasses import dataclass
import math
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import mkdtemp

_DEFAULT_STARTUP_TIMEOUT_SEC = 15.0
_DEFAULT_EXEC_TIMEOUT_SEC = 30.0


class LibreOfficeUnavailableError(RuntimeError):
    """Raised when the LibreOffice runtime is not available."""


@dataclass(frozen=True)
class LibreOfficeSessionConfig:
    soffice_path: Path
    startup_timeout_sec: float
    exec_timeout_sec: float
    profile_root: Path | None


class LibreOfficeSession:
    """Best-effort runtime guard for LibreOffice-backed extraction."""

    def __init__(self, config: LibreOfficeSessionConfig) -> None:
        self.config = config
        self._temp_profile_dir: Path | None = None

    @classmethod
    def from_env(cls) -> LibreOfficeSession:
        """Build a session from ExStruct environment variables."""
        raw_path = os.getenv("EXSTRUCT_LIBREOFFICE_PATH")
        resolved = Path(raw_path) if raw_path else _which_soffice()
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
        if not self.config.soffice_path.exists():
            raise LibreOfficeUnavailableError(
                f"LibreOffice runtime is unavailable: '{self.config.soffice_path}' was not found."
            )
        profile_root = self.config.profile_root
        if profile_root is not None:
            profile_root.mkdir(parents=True, exist_ok=True)
            self._temp_profile_dir = Path(
                mkdtemp(prefix="exstruct-lo-", dir=str(profile_root))
            )
        else:
            self._temp_profile_dir = Path(mkdtemp(prefix="exstruct-lo-"))
        try:
            subprocess.run(
                [str(self.config.soffice_path), "--version"],
                capture_output=True,
                check=True,
                text=True,
                timeout=self.config.startup_timeout_sec,
            )
        except Exception as exc:
            if self._temp_profile_dir is not None:
                shutil.rmtree(self._temp_profile_dir, ignore_errors=True)
                self._temp_profile_dir = None
            if isinstance(exc, FileNotFoundError):
                raise LibreOfficeUnavailableError(
                    f"LibreOffice runtime is unavailable: '{self.config.soffice_path}' could not be executed."
                ) from exc
            if isinstance(exc, subprocess.TimeoutExpired):
                raise LibreOfficeUnavailableError(
                    "LibreOffice runtime is unavailable: soffice version probe timed out."
                ) from exc
            if isinstance(exc, subprocess.CalledProcessError):
                raise LibreOfficeUnavailableError(
                    "LibreOffice runtime is unavailable: soffice version probe failed."
                ) from exc
            raise
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        _ = exc_type
        _ = exc
        _ = tb
        if self._temp_profile_dir is not None:
            shutil.rmtree(self._temp_profile_dir, ignore_errors=True)
            self._temp_profile_dir = None

    def load_workbook(self, file_path: Path) -> object:
        """Return a lightweight workbook token for future subprocess integration."""
        return {"file_path": str(file_path.resolve())}

    def close_workbook(self, workbook: object) -> None:
        """Close a workbook token returned by ``load_workbook``."""
        _ = workbook


def _which_soffice() -> Path | None:
    for candidate in ("soffice", "soffice.exe"):
        resolved = shutil.which(candidate)
        if resolved:
            return Path(resolved)
    return None


def _get_timeout_from_env(name: str, *, default: float) -> float:
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
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return None
    return Path(raw)
