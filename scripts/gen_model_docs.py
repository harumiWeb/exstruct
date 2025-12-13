from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass
import importlib
import inspect
from pathlib import Path
import sys
from types import UnionType
from typing import Union, get_args, get_origin

from pydantic import BaseModel


@dataclass
class FieldDoc:
    """Structured information about a Pydantic field."""

    name: str
    type_repr: str
    required: bool
    default: str
    description: str


@dataclass
class ModelDoc:
    """Structured documentation for a Pydantic model."""

    name: str
    qualname: str
    docstring: str | None
    fields: list[FieldDoc]


def project_root() -> Path:
    """
    Return the repository root based on this file location.

    Returns:
        Path to the repository root.
    """
    return Path(__file__).resolve().parent.parent


def _format_default(default: object, *, has_factory: bool) -> str:
    """
    Render a default value for documentation.

    Args:
        default: Default value from a field.
        has_factory: Whether the field uses a default_factory.

    Returns:
        Human-friendly default representation.
    """
    if has_factory:
        return "<factory>"
    if default is None:
        return "None"
    if isinstance(default, str | int | float | bool):
        return repr(default)
    return type(default).__name__


def _format_annotation(annotation: object) -> str:
    """
    Render a type annotation into a concise string.

    Args:
        annotation: Type annotation from a field.

    Returns:
        String representation of the annotation.
    """
    if isinstance(annotation, str):
        return annotation

    origin = get_origin(annotation)
    args = get_args(annotation)

    if origin in {Union, UnionType}:
        return " | ".join(sorted({_format_annotation(arg) for arg in args}))

    if origin is tuple and len(args) == 2 and args[1] is Ellipsis:
        return f"tuple[{_format_annotation(args[0])}, ...]"

    if str(origin).endswith("Literal"):
        return f"Literal[{', '.join(repr(arg) for arg in args)}]"

    if origin:
        origin_name = getattr(origin, "__name__", str(origin))
        formatted_args = ", ".join(_format_annotation(arg) for arg in args)
        return f"{origin_name}[{formatted_args}]"

    if hasattr(annotation, "__name__"):
        return annotation.__name__  # type: ignore[no-any-return]

    return str(annotation)


def _collect_models(module_names: Sequence[str]) -> list[type[BaseModel]]:
    """
    Import modules and collect local BaseModel subclasses.

    Args:
        module_names: Sequence of module import paths to inspect.

    Returns:
        Sorted list of BaseModel subclasses found in the modules.
    """
    models: list[type[BaseModel]] = []
    for module_name in module_names:
        if module_name == "exstruct.models":
            module = importlib.import_module("exstruct.models")
        elif module_name == "exstruct.engine":
            module = importlib.import_module("exstruct.engine")
        else:
            msg = f"Unsafe module import blocked: {module_name}"
            raise ValueError(msg)
        for _, cls in inspect.getmembers(module, inspect.isclass):
            if not issubclass(cls, BaseModel):
                continue
            if cls.__module__ != module.__name__:
                continue
            cls.model_rebuild()
            models.append(cls)
    return sorted(models, key=lambda cls: cls.__name__)


def _field_docs(model: type[BaseModel]) -> list[FieldDoc]:
    """
    Build field documentation entries for a model.

    Args:
        model: Pydantic BaseModel subclass.

    Returns:
        List of field documentation entries.
    """
    docs: list[FieldDoc] = []
    for name, field in sorted(model.model_fields.items()):
        type_repr = _format_annotation(field.annotation)
        required = field.is_required()
        has_factory = field.default_factory is not None
        default = (
            "" if required else _format_default(field.default, has_factory=has_factory)
        )
        description = (
            (field.description or "").replace("|", "\\|").replace("\n", "<br>")
        )
        docs.append(
            FieldDoc(
                name=name,
                type_repr=type_repr,
                required=required,
                default=default,
                description=description,
            )
        )
    return docs


def _escape_cell(text: str) -> str:
    """
    Escape characters that would break markdown tables.

    Args:
        text: Raw cell text.

    Returns:
        Escaped cell text.
    """
    return text.replace("|", "\\|")


def _model_docs(models: Iterable[type[BaseModel]]) -> list[ModelDoc]:
    """
    Convert model classes into documentation structures.

    Args:
        models: Iterable of BaseModel subclasses.

    Returns:
        List of model documentation entries.
    """
    base_doc = inspect.getdoc(BaseModel)
    docs: list[ModelDoc] = []
    for model in models:
        docstring = inspect.getdoc(model)
        if docstring == base_doc:
            docstring = None
        fields = _field_docs(model)
        docs.append(
            ModelDoc(
                name=model.__name__,
                qualname=model.__qualname__,
                docstring=docstring,
                fields=fields,
            )
        )
    return docs


def _render_markdown(model_docs: Sequence[ModelDoc]) -> str:
    """
    Render collected model docs into a markdown string.

    Args:
        model_docs: Documentation entries to render.

    Returns:
        Markdown content.
    """
    lines: list[str] = [
        "<!-- Auto-generated by scripts/gen_model_docs.py; do not edit by hand. -->",
        "# Data Models",
        "",
    ]

    for model in model_docs:
        lines.append(f"## {model.name}")
        if model.docstring:
            lines.append(model.docstring)
        lines.append("")
        lines.append("| Field | Type | Required | Default | Description |")
        lines.append("| --- | --- | --- | --- | --- |")
        for field in model.fields:
            required = "Yes" if field.required else "No"
            default = _escape_cell(field.default) if field.default else "-"
            description = _escape_cell(field.description or "-")
            type_repr = _escape_cell(field.type_repr)
            lines.append(
                f"| `{field.name}` | `{type_repr}` | {required} | {default} | {description} |"
            )
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main() -> None:
    """
    Entry point: generate docs/generated/models.md from Pydantic models.
    """
    root = project_root()
    sys.path.insert(0, str(root))

    module_names = [
        "exstruct.models",
        "exstruct.engine",
    ]
    models = _collect_models(module_names)
    model_docs = _model_docs(models)
    output = _render_markdown(model_docs)

    output_dir = root / "docs" / "generated"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "models.md"
    output_path.write_text(output, encoding="utf-8")


if __name__ == "__main__":
    main()
