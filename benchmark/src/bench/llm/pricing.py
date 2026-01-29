from __future__ import annotations

# gpt-4o pricing: per 1M tokens
# Input $2.50 / 1M, Output $10.00 / 1M (cached input not used here)
# Source: model compare page
# https://platform.openai.com/docs/models/compare?model=gpt-4o
# (You will cite this in README/report; code keeps constants.)
GPT4O_INPUT_PER_1M = 2.50
GPT4O_OUTPUT_PER_1M = 10.00

_PRICING_PER_1M: dict[str, tuple[float, float]] = {
    "gpt-4o": (GPT4O_INPUT_PER_1M, GPT4O_OUTPUT_PER_1M),
}


def estimate_cost_usd(model: str, input_tokens: int, output_tokens: int) -> float:
    """Estimate USD cost for a model run when pricing is known."""
    pricing = _PRICING_PER_1M.get(model)
    if pricing is None:
        # Pricing unknown; keep run going and report 0.0 cost.
        return 0.0
    input_per_1m, output_per_1m = pricing
    return (input_tokens / 1_000_000) * input_per_1m + (
        output_tokens / 1_000_000
    ) * output_per_1m
