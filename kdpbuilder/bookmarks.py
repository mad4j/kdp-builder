from __future__ import annotations

import re


def sanitize_bookmark_name(text: str) -> str:
    """Convert heading text to a valid Word bookmark name."""
    sanitized = re.sub(r"[^a-zA-Z0-9\s]", "", text)
    sanitized = sanitized.replace(" ", "_").strip("_")

    if not sanitized:
        return "bookmark"

    if not sanitized[0].isalpha():
        sanitized = "h_" + sanitized

    sanitized = sanitized[:40]
    return sanitized.lower()
