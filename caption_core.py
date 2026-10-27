"""Small dependency-free helpers for splitting PowerPoint speaker notes."""
import re


def split_captions(notes: str | None) -> list[str]:
    """Respect '//' segments; otherwise split by sentence boundary/newline.

    The original presentation notes remain untouched. This is intentionally
    conservative: presenters can edit the saved caption cues themselves.
    """
    text = (notes or "").replace("\r", "\n")
    if "//" in text:
        segments = re.split(r"\s*//\s*", text)
    else:
        segments = re.split(r"(?<=[.!?。！？])\s+|\n+", text)
    return [segment.strip() for segment in segments if segment.strip()]