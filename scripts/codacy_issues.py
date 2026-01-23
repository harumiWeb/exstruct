#!/usr/bin/env python3

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
import re
import subprocess  # nosec B404 - used for fixed git commands only
import sys
from typing import Any, cast
import urllib.parse
import urllib.request

# ================================
# Config
# ================================
BASE = "https://api.codacy.com/api/v3"
BASE_URL = urllib.parse.urlparse(BASE)
BASE_PATH = BASE_URL.path.rstrip("/")  # "/api/v3"


def get_token() -> str:
    """Return the Codacy API token or raise if missing.

    Returns:
        Codacy API token string from the environment.

    Raises:
        ValueError: If CODACY_API_TOKEN is not set.
    """
    token = os.environ.get("CODACY_API_TOKEN")
    if token is None:
        raise ValueError("CODACY_API_TOKEN is not set")
    return token


# ================================
# Utilities
# ================================
LEVELS = ["Error", "High", "Warning", "Info"]


def get_level_priority(level: str | None) -> int | None:
    """Convert a severity level name to a priority number.

    Args:
        level: Severity level string.

    Returns:
        Priority number or None if unknown.
    """
    if level == "Error":
        return 4
    if level == "High":
        return 3
    if level == "Warning":
        return 2
    if level == "Info":
        return 1
    return None


def normalize_provider(value: str) -> str | None:
    """
    Normalize a provider identifier to a supported short code.
    
    Parameters:
        value (str): Provider identifier to normalize (expected 'gh', 'gl', or 'bb').
    
    Returns:
        str | None: The provider code ('gh', 'gl', or 'bb') if valid, `None` otherwise.
    """
    return value if value in ("gh", "gl", "bb") else None


def assert_valid_segment(name: str, value: str, pattern: re.Pattern[str]) -> str:
    """Validate an identifier segment against a regex.

    Args:
        name: Segment name for error reporting.
        value: Segment value.
        pattern: Compiled regex pattern for allowed values.

    Returns:
        The validated value.

    Raises:
        ValueError: If the value is empty or invalid.
    """
    if (not value) or (pattern.match(value) is None):
        raise ValueError(f"Invalid {name}: {value}")
    return value


def assert_valid_choice(name: str, value: str, choices: list[str]) -> str:
    """Validate that a value is in a list of choices.

    Args:
        name: Parameter name for error reporting.
        value: Input value.
        choices: Allowed values.

    Returns:
        The validated value.

    Raises:
        ValueError: If the value is not allowed.
    """
    if value not in choices:
        raise ValueError(f"Invalid {name}: {value}. Valid values: {', '.join(choices)}")
    return value


def encode_segment(value: str) -> str:
    """
    URL-encode a URL path segment so it is safe for inclusion in a path.
    
    Returns:
        encoded (str): The percent-encoded representation of the input string.
    """
    return urllib.parse.quote(value, safe="")


def build_codacy_url(pathname: str, query: dict[str, str] | None = None) -> str:
    """
    Constructs a full Codacy API URL using the configured base origin and base path.
    
    Parameters:
        pathname (str): Pathname to append to the base path (should begin with a forward slash).
        query (dict[str, str] | None): Optional mapping of query parameter names to values; values are URL-encoded.
    
    Returns:
        url (str): The complete URL including query string if `query` is provided.
    """
    # Ensure we keep origin and base path
    url = f"{BASE_URL.scheme}://{BASE_URL.netloc}{BASE_PATH}{pathname}"
    if query:
        url = f"{url}?{urllib.parse.urlencode(query)}"
    return url


def assert_codacy_url(url: str) -> str:
    """
    Validate that `url` targets the configured Codacy API origin and begins with the `/analysis/` path.
    
    Parameters:
        url (str): The full URL to validate.
    
    Returns:
        str: The original URL when it is confirmed to target the configured Codacy API origin and start with the `/analysis/` path.
    
    Raises:
        ValueError: If the URL does not use the configured Codacy API origin or does not start with the expected `/analysis/` path.
    """
    # Basic safety: must be same origin and start with /api/v3/analysis/
    parsed = urllib.parse.urlparse(url)
    expected_origin = f"{BASE_URL.scheme}://{BASE_URL.netloc}"
    origin = f"{parsed.scheme}://{parsed.netloc}"
    expected_prefix = f"{BASE_PATH}/analysis/"
    if origin != expected_origin or not parsed.path.startswith(expected_prefix):
        raise ValueError(f"Invalid URL: {url}")
    return url


def build_repo_issues_url(provider: str, org: str, repo: str, limit: int) -> str:
    """
    Constructs the Codacy API URL to search repository issues for a given provider, organization, repository, and result limit.
    
    Parameters:
        provider (str): Provider code (e.g., "gh", "gl", "bb").
        org (str): Organization or owner name.
        repo (str): Repository name.
        limit (int): Maximum number of results to request.
    
    Returns:
        str: A Codacy API URL for the repository issues search endpoint with the `limit` query parameter set.
    """
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/issues/search",
        query={"limit": str(limit)},
    )


def build_pr_issues_url(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str
) -> str:
    """
    Constructs the Codacy API URL for fetching issues of a pull request.
    
    Parameters:
        provider (str): Provider code (e.g., "gh", "gl", "bb").
        org (str): Organization or owner name.
        repo (str): Repository name.
        pr (str): Pull request identifier.
        limit (int): Maximum number of issues to request.
        status (str): Issue status filter (e.g., "all", "open", "closed").
    
    Returns:
        str: The Codacy API URL for the pull-request issues endpoint including `status` and `limit` query parameters.
    """
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/pull-requests/{encode_segment(pr)}/issues",
        query={"status": status, "limit": str(limit)},
    )


def get_git_origin_url() -> str | None:
    """
    Get the Git remote "origin" URL for the current repository, or None when it cannot be determined.
    
    Returns:
        origin_url (str | None): The remote URL configured for 'origin' if the current directory is inside a Git work tree and the origin URL is available; `None` if not inside a Git repository, if the origin is not set, or on error.
    """
    # git repo check
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--is-inside-work-tree"],
            capture_output=True,
            text=True,
            check=False,
        )  # nosec B603 - fixed git command without user input
        if result.returncode != 0 or not result.stdout.strip():
            return None
        result = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            capture_output=True,
            text=True,
            check=False,
        )  # nosec B603 - fixed git command without user input
        if result.returncode != 0:
            return None
        return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return None


@dataclass
class GitRemoteInfo:
    """Parsed git remote information."""

    provider: str
    org: str
    repo: str


def parse_git_remote(url: str) -> GitRemoteInfo | None:
    """
    Extract provider, organization, and repository from a Git remote URL.
    
    Accepts HTTPS (https://host/org/repo[.git]) and SSH (git@host:org/repo[.git]) remote formats.
    Provider is one of: "gh" for GitHub, "gl" for GitLab, "bb" for Bitbucket, or "unknown" for other hosts.
    
    Parameters:
        url (str): Git remote URL to parse.
    
    Returns:
        GitRemoteInfo | None: Parsed GitRemoteInfo with fields `provider`, `org`, and `repo`, or `None` if the URL could not be parsed.
    """
    # HTTPS
    m = re.match(r"^https?://([^/]+)/([^/]+)/([^/]+?)(?:\.git)?$", url)
    # SSH
    if not m:
        m = re.match(r"^git@([^:]+):([^/]+)/([^/]+?)(?:\.git)?$", url)

    if not m:
        return None

    host, org, repo = m.group(1), m.group(2), m.group(3)

    def is_same_or_subdomain(hostname: str, base_domain: str) -> bool:
        """
        Check whether a hostname is equal to a base domain or is a subdomain of that base domain.
        
        Parameters:
            hostname (str): Hostname to test (e.g., "api.example.com").
            base_domain (str): Base domain to compare against (e.g., "example.com").
        
        Returns:
            `true` if `hostname` equals `base_domain` or ends with `.` followed by `base_domain`, `false` otherwise.
        """
        return hostname == base_domain or hostname.endswith("." + base_domain)

    if is_same_or_subdomain(host, "github.com"):
        provider = "gh"
    elif is_same_or_subdomain(host, "gitlab.com"):
        provider = "gl"
    elif is_same_or_subdomain(host, "bitbucket.org"):
        provider = "bb"
    else:
        provider = "unknown"

    return GitRemoteInfo(provider=provider, org=org, repo=repo)


def fetch_json(
    url: str, method: str = "GET", body: dict[str, Any] | None = None
) -> dict[str, Any]:
    """
    Fetch and return a JSON object from a validated Codacy API URL.
    
    Parameters:
        url (str): Codacy API URL; must target the configured Codacy origin and start with the /analysis/ path.
        method (str): HTTP method to use (e.g., "GET", "POST").
        body (dict[str, Any] | None): Optional JSON body for non-GET requests.
    
    Returns:
        dict[str, Any]: The parsed JSON response as a dictionary.
    
    Raises:
        RuntimeError: On HTTP errors, network errors, invalid JSON, or when the JSON root value is not an object.
    """
    safe_url = assert_codacy_url(url)

    headers = {
        "Accept": "application/json",
        "api-token": get_token(),
    }

    data: bytes | None = None
    if body is not None and method.upper() != "GET":
        payload = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
        headers["Content-Length"] = str(len(payload))
        data = payload

    req = urllib.request.Request(
        safe_url, method=method.upper(), headers=headers, data=data
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as res:  # nosec B310 - validated https origin
            raw = res.read().decode("utf-8", errors="replace")
            try:
                parsed = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise RuntimeError("Invalid JSON response") from exc
            if not isinstance(parsed, dict):
                raise RuntimeError("Invalid JSON response")
            return cast(dict[str, Any], parsed)
    except urllib.error.HTTPError as e:
        # include response body if possible
        try:
            body_text = e.read().decode("utf-8", errors="replace")
        except Exception:
            body_text = ""
        raise RuntimeError(f"HTTP {e.code}: {body_text or str(e)}") from None
    except urllib.error.URLError as e:
        raise RuntimeError(str(e)) from None


# ================================
# API
# ================================
def fetch_repo_issues(provider: str, org: str, repo: str, limit: int) -> dict[str, Any]:
    """
    Request Codacy for issues belonging to a repository.
    
    Parameters:
        provider (str): Provider code ('gh', 'gl', 'bb') indicating GitHub, GitLab, or Bitbucket.
        org (str): Organization or owner name.
        repo (str): Repository name.
        limit (int): Maximum number of issues to return.
    
    Returns:
        dict[str, Any]: Parsed JSON response from the Codacy API containing issue data.
    """
    url = build_repo_issues_url(provider, org, repo, limit)
    return fetch_json(url, method="POST", body={})


def fetch_pr_issues(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str = "all"
) -> dict[str, Any]:
    """
    Retrieve Codacy issues for a specific pull request.
    
    Parameters:
        provider (str): Provider code ("gh", "gl", "bb").
        org (str): Organization or user name.
        repo (str): Repository name.
        pr (str): Pull request number or identifier.
        limit (int): Maximum number of issues to request.
        status (str): Issue status filter (for example "all", "open", "closed").
    
    Returns:
        dict: Parsed JSON response from the Codacy API.
    """
    url = build_pr_issues_url(provider, org, repo, pr, limit, status)
    return fetch_json(url, method="GET")


# ================================
# AI Output Formatter
# ================================
def format_for_ai(raw_issues: list[dict[str, Any]], min_level: str) -> list[str]:
    """
    Format Codacy issue records into compact AI-friendly lines filtered by minimum severity.
    
    Each returned string has the form:
    "<level> | <file_path>:<line_no> | <rule> | <category> | <message>".
    
    Parameters:
        raw_issues: List of issue objects returned by the Codacy API (each item may be an issue or contain a `commitIssue` key).
        min_level: Minimum severity level to include; must be one of the values in LEVELS.
    
    Returns:
        A list of formatted issue strings matching the format above, including only issues whose severity is at or above `min_level`.
    
    Raises:
        ValueError: If `min_level` is not a valid severity level.
    """
    min_priority = get_level_priority(min_level)
    if min_priority is None:
        raise ValueError(
            f"Invalid min_level: {min_level}. Valid values: {', '.join(LEVELS)}"
        )

    out: list[str] = []

    for item in raw_issues:
        issue = item.get("commitIssue") or item

        pattern_info = issue.get("patternInfo") or {}
        level = pattern_info.get("level")
        prio = get_level_priority(level)
        if prio is None or prio < min_priority:
            continue

        file_path = issue.get("filePath")
        line_no = issue.get("lineNumber")
        rule = pattern_info.get("id")
        category = pattern_info.get("category")
        message = issue.get("message")

        out.append(f"{level} | {file_path}:{line_no} | {rule} | {category} | {message}")

    return out


# ================================
# CLI
# ================================
def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument("org", nargs="?", default=None)
    p.add_argument("repo", nargs="?", default=None)
    p.add_argument("--pr", dest="pr", default=None)
    p.add_argument("--min-level", dest="min_level", default="Info", choices=LEVELS)
    p.add_argument("--provider", dest="provider", default=None)
    p.add_argument("--help", action="help", help="Show this help message and exit")
    return p.parse_args(argv)


def apply_git_defaults(args: argparse.Namespace) -> None:
    """Populate missing org/repo/provider from git origin when possible."""
    if args.org and args.repo:
        return
    origin_url = get_git_origin_url()
    if not origin_url:
        return
    parsed = parse_git_remote(origin_url)
    if not parsed:
        return
    if args.provider is None:
        args.provider = parsed.provider
    if args.org is None:
        args.org = parsed.org
    if args.repo is None:
        args.repo = parsed.repo


def resolve_segments(args: argparse.Namespace) -> tuple[str, str, str | None]:
    """
    Validate CLI org, repo, and optional pr segments and return them.
    
    Parameters:
        args (argparse.Namespace): Parsed CLI arguments with attributes `org`, `repo`, and optional `pr`.
    
    Returns:
        tuple[str, str, str | None]: A tuple (org, repo, pr) where `pr` is None if not supplied.
    
    Raises:
        ValueError: If any segment is empty or contains invalid characters.
    """
    segment_pattern = re.compile(r"^[A-Za-z0-9_.-]+$")
    org = assert_valid_segment("org", args.org, segment_pattern)
    repo = assert_valid_segment("repo", args.repo, segment_pattern)
    pr = args.pr
    if pr is not None:
        pr = assert_valid_segment("pr", pr, re.compile(r"^[0-9]+$"))
    return org, repo, pr


def build_payload(
    *,
    pr: str | None,
    org: str,
    repo: str,
    min_level: str,
    issues: list[str],
) -> dict[str, object]:
    """
    Create a JSON-serializable payload describing the fetched issues and their scope.
    
    The returned dictionary contains:
    - scope: "pull_request" when `pr` is set, otherwise "repository".
    - organization: organization/owner name.
    - repository: repository name.
    - pullRequest: pull request identifier string when present, otherwise `None`.
    - minLevel: the minimum severity level used to filter issues.
    - total: the number of issues in `issues`.
    - issues: list of formatted issue strings.
    
    Returns:
        dict[str, object]: Payload ready for JSON serialization with the keys described above.
    """
    return {
        "scope": "pull_request" if pr else "repository",
        "organization": org,
        "repository": repo,
        "pullRequest": pr if pr else None,
        "minLevel": min_level,
        "total": len(issues),
        "issues": issues,
    }


def main() -> int:
    """
    Run the CLI: parse arguments, fetch Codacy issues (repository or pull request), format them for AI consumption, and write a JSON payload to stdout.
    
    Writes error messages to stderr when validation or fetching fails and prints the final JSON payload to stdout.
    
    Returns:
        int: 0 on success, 1 on error.
    """
    args = parse_args(sys.argv[1:])

    # --- Git auto-detect ---
    apply_git_defaults(args)

    if args.provider is None:
        args.provider = "gh"

    provider = normalize_provider(args.provider)
    if not provider:
        print("Invalid --provider: use gh, gl, or bb", file=sys.stderr)
        return 1

    if not args.org or not args.repo:
        print(
            "Usage:\n"
            "  python codacy_issues.py ORG REPO [--pr NUMBER] [--min-level Error|High|Warning|Info] [--provider gh|gl|bb]",
            file=sys.stderr,
        )
        return 1

    try:
        org, repo, pr = resolve_segments(args)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    status = "all"
    limit = 100

    result = (
        fetch_pr_issues(
            provider=provider, org=org, repo=repo, pr=pr, limit=limit, status=status
        )
        if pr
        else fetch_repo_issues(provider=provider, org=org, repo=repo, limit=limit)
    )

    issues = result.get("data") or []
    try:
        formatted = format_for_ai(issues, args.min_level)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    payload = build_payload(
        pr=pr, org=org, repo=repo, min_level=args.min_level, issues=formatted
    )

    sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(1) from e