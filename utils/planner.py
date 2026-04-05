"""
Slide planner – provides a fallback planning function that runs purely
from the agent's knowledge (no external calls) when the MCP plan_slides
tool is unavailable.
"""

import json
import re


def parse_num_slides(prompt: str, default: int = 5) -> int:
    """Extract requested slide count from a natural-language prompt."""
    match = re.search(r"\b(\d+)[- ]?slide", prompt, re.IGNORECASE)
    if match:
        n = int(match.group(1))
        return max(2, min(n, 15))
    return default


def extract_topic(prompt: str) -> str:
    """Best-effort extraction of the core topic from a prompt."""
    # Remove common preamble phrases
    clean = re.sub(
        r"(create|make|build|generate|prepare|produce)\s+(a|an|me|the)?\s*"
        r"(\d+[- ]?slide\s+)?(presentation|ppt|powerpoint|deck|slideshow)\s*(on|about|for|regarding)?\s*",
        "",
        prompt,
        flags=re.IGNORECASE,
    ).strip(" .,")
    return clean or prompt


def build_fallback_plan(topic: str, num_slides: int) -> list[dict]:
    """
    Generate a generic slide outline when AI planning is unavailable.
    Always produces exactly `num_slides` slide dicts.
    """
    plan = [{"title": topic, "bullets": []}]
    sections = [
        ("Introduction", [f"What is {topic}?", "Why it matters", "Overview of key concepts"]),
        ("Background", ["Historical context", "Foundational principles", "Key terminology"]),
        ("Core Concepts", ["Concept 1 – definition and significance",
                           "Concept 2 – how it works",
                           "Concept 3 – real-world examples"]),
        ("Deep Dive", ["Detailed analysis", "Supporting evidence", "Case studies"]),
        ("Applications", ["Practical use cases", "Industry impact", "Future potential"]),
        ("Challenges", ["Common misconceptions", "Known limitations", "Open research questions"]),
        ("Current Trends", ["Recent developments", "Emerging technologies", "Expert perspectives"]),
        ("Future Outlook", ["Predictions", "Opportunities ahead", "What to watch for"]),
        ("Summary", ["Key takeaways", "Recap of main points", "Action items"]),
        ("Q&A / References", ["Further reading", "Sources", "Thank you"]),
    ]
    for i in range(1, num_slides):
        if i - 1 < len(sections):
            title, bullets = sections[i - 1]
        else:
            title, bullets = f"Section {i}", ["Key point A", "Key point B", "Key point C"]
        plan.append({"title": title, "bullets": bullets})
    return plan[:num_slides]