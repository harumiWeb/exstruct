from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI
from pydantic import BaseModel

from ..paths import ROOT
from .pricing import estimate_cost_usd


class LLMResult(BaseModel):
    """Structured response data from the LLM call."""

    text: str
    input_tokens: int
    output_tokens: int
    cost_usd: float
    raw: dict[str, Any]


def _png_to_data_url(png_path: Path) -> str:
    """Encode a PNG image as a data URL.

    Args:
        png_path: PNG file path.

    Returns:
        Base64 data URL string.
    """
    b = png_path.read_bytes()
    b64 = base64.b64encode(b).decode("ascii")
    return f"data:image/png;base64,{b64}"


def _extract_usage_tokens(usage: object | None) -> tuple[int, int]:
    """Extract input/output tokens from the OpenAI usage payload.

    Args:
        usage: Usage payload from the OpenAI SDK (object or dict).

    Returns:
        Tuple of (input_tokens, output_tokens).
    """
    if usage is None:
        return 0, 0
    if isinstance(usage, dict):
        return int(usage.get("input_tokens", 0)), int(usage.get("output_tokens", 0))
    input_tokens = int(getattr(usage, "input_tokens", 0))
    output_tokens = int(getattr(usage, "output_tokens", 0))
    return input_tokens, output_tokens


class OpenAIResponsesClient:
    """Thin wrapper around the OpenAI Responses API for this benchmark."""

    def __init__(self) -> None:
        load_dotenv(dotenv_path=ROOT / ".env")
        if not os.getenv("OPENAI_API_KEY"):
            raise RuntimeError(
                "OPENAI_API_KEY is not set. Add it to .env or your environment."
            )
        self.client = OpenAI()

    def ask_text(
        self, *, model: str, question: str, context_text: str, temperature: float
    ) -> LLMResult:
        """Call Responses API with text-only input.

        Args:
            model: OpenAI model name (e.g., "gpt-4o").
            question: User question to answer.
            context_text: Extracted context text from the workbook.
            temperature: Sampling temperature for the response.

        Returns:
            LLMResult containing the model output and usage metadata.
        """
        resp = self.client.responses.create(
            model=model,
            temperature=temperature,
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "You are a strict JSON extraction engine. Output JSON only.",
                        },
                        {"type": "input_text", "text": f"[QUESTION]\n{question}"},
                        {"type": "input_text", "text": f"[CONTEXT]\n{context_text}"},
                    ],
                }
            ],
        )

        text = resp.output_text  # SDK helper
        usage = getattr(resp, "usage", None)
        in_tok, out_tok = _extract_usage_tokens(usage)
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )

    def ask_images(
        self, *, model: str, question: str, image_paths: list[Path], temperature: float
    ) -> LLMResult:
        """Call Responses API with image + text input.

        Args:
            model: OpenAI model name (e.g., "gpt-4o").
            question: User question to answer.
            image_paths: PNG image paths to include as vision input.
            temperature: Sampling temperature for the response.

        Returns:
            LLMResult containing the model output and usage metadata.
        """
        content: list[dict[str, Any]] = [
            {
                "type": "input_text",
                "text": "You are a strict JSON extraction engine. Output JSON only.",
            },
            {"type": "input_text", "text": f"[QUESTION]\n{question}"},
        ]
        for p in image_paths:
            content.append({"type": "input_image", "image_url": _png_to_data_url(p)})

        resp = self.client.responses.create(
            model=model,
            temperature=temperature,
            input=[{"role": "user", "content": content}],
        )

        text = resp.output_text
        usage = getattr(resp, "usage", None)
        in_tok, out_tok = _extract_usage_tokens(usage)
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )

    def ask_markdown(
        self, *, model: str, json_text: str, temperature: float
    ) -> LLMResult:
        """Call Responses API to convert JSON into Markdown.

        Args:
            model: OpenAI model name (e.g., "gpt-4o").
            json_text: JSON payload to convert to Markdown.
            temperature: Sampling temperature for the response.

        Returns:
            LLMResult containing the model output and usage metadata.
        """
        instructions = (
            "You are a strict Markdown formatter. Output Markdown only.\n"
            "Rules:\n"
            "- Use '## <key>' for top-level keys.\n"
            "- For lists of scalars, use bullet lists.\n"
            "- For lists of objects, use Markdown tables with columns in key order.\n"
            "- For nested objects or lists inside table cells, use compact JSON.\n"
        )
        resp = self.client.responses.create(
            model=model,
            temperature=temperature,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": instructions},
                        {"type": "input_text", "text": f"[JSON]\n{json_text}"},
                    ],
                }
            ],
        )

        text = resp.output_text
        usage = getattr(resp, "usage", None)
        in_tok, out_tok = _extract_usage_tokens(usage)
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )

    def ask_markdown_from_text(
        self, *, model: str, context_text: str, temperature: float
    ) -> LLMResult:
        """Call Responses API to convert raw text into Markdown.

        Args:
            model: OpenAI model name (e.g., "gpt-4o").
            context_text: Extracted document text to format.
            temperature: Sampling temperature for the response.

        Returns:
            LLMResult containing the model output and usage metadata.
        """
        instructions = (
            "You are a strict Markdown formatter. Output Markdown only.\n"
            "Rules:\n"
            "- Preserve all content from the input.\n"
            "- Use headings and lists when they are clearly implied.\n"
            "- Use tables when a row/column structure is evident.\n"
            "- Do not add or invent information.\n"
        )
        resp = self.client.responses.create(
            model=model,
            temperature=temperature,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": instructions},
                        {"type": "input_text", "text": f"[TEXT]\n{context_text}"},
                    ],
                }
            ],
        )

        text = resp.output_text
        usage = getattr(resp, "usage", None)
        in_tok, out_tok = _extract_usage_tokens(usage)
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )

    def ask_markdown_images(
        self, *, model: str, image_paths: list[Path], temperature: float
    ) -> LLMResult:
        """Call Responses API to convert images into Markdown.

        Args:
            model: OpenAI model name (e.g., "gpt-4o").
            image_paths: PNG image paths to include as vision input.
            temperature: Sampling temperature for the response.

        Returns:
            LLMResult containing the model output and usage metadata.
        """
        instructions = (
            "You are a strict Markdown formatter. Output Markdown only.\n"
            "Rules:\n"
            "- Preserve all visible content from the images.\n"
            "- Use headings and lists when they are clearly implied.\n"
            "- Use tables when a row/column structure is evident.\n"
            "- Do not add or invent information.\n"
        )
        content: list[dict[str, Any]] = [
            {"type": "input_text", "text": instructions},
        ]
        for p in image_paths:
            content.append({"type": "input_image", "image_url": _png_to_data_url(p)})

        resp = self.client.responses.create(
            model=model,
            temperature=temperature,
            input=[{"role": "user", "content": content}],
        )

        text = resp.output_text
        usage = getattr(resp, "usage", None)
        in_tok, out_tok = _extract_usage_tokens(usage)
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )
