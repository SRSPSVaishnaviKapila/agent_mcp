"""
Agent core – agentic loop using MCPClient as a proper async context manager
to avoid ExceptionGroup / TaskGroup crashes on Windows.
"""

import os
from typing import Callable

from agent.mcp_client import MCPClient
from agent.logger import StepLog
from utils.planner import extract_topic, parse_num_slides, build_fallback_plan


async def run_agent(
    user_request: str,
    output_dir: str = "outputs",
    progress_cb: Callable[[str], None] | None = None,
) -> dict:

    log = StepLog()

    def emit(msg: str):
        if progress_cb:
            progress_cb(msg)

    # ── Parse request ──────────────────────────────────────────────────────────
    topic      = extract_topic(user_request)
    num_slides = parse_num_slides(user_request, default=5)
    filename   = "ai.pptx"
    include_img = os.getenv("INCLUDE_IMAGE_SLIDE", "1") == "1"

    emit(f"📋 Topic: **{topic}** | Slides: **{num_slides}** | Image slide: **{include_img}**")
    log.info("Parse", f"topic='{topic}', num_slides={num_slides}")

    os.environ["PPT_OUTPUT_DIR"] = output_dir

    plan: list[dict] = []
    slides_added = 0

    # ── All MCP work happens inside the context manager ────────────────────────
    # This ensures proper TaskGroup cleanup even on Windows / Python 3.11+
    try:
        emit("🔌 Connecting to MCP server…")

        async with MCPClient() as client:
            tools = await client.list_tools()
            emit(f"✅ Connected. Tools: `{', '.join(tools)}`")
            log.ok("MCP Connect", f"tools={tools}")

            # ── Step 1: Plan ───────────────────────────────────────────────────
            emit("🧠 Step 1 – Planning slides with HuggingFace…")
            try:
                plan = await client.plan_slides(topic, num_slides)
                log.ok("Plan", f"{len(plan)} slides planned")
                emit(f"📝 Plan ready ({len(plan)} slides)")
            except Exception as e:
                log.error("Plan", str(e))
                plan = build_fallback_plan(topic, num_slides)
                emit(f"⚠️ HF planner failed — using static fallback ({len(plan)} slides)")

            # ── Step 2: Create presentation ────────────────────────────────────
            emit("🎨 Step 2 – Creating presentation…")
            res = await client.create_presentation()
            log.ok("Create", res)
            emit(f"✅ {res}")

            # ── Step 3: Build slides ───────────────────────────────────────────
            emit("🖼️ Step 3 – Building slides…")
            img_slide_index = max(1, len(plan) // 2) if include_img else -1

            for i, slide_def in enumerate(plan):
                # Normalise
                if isinstance(slide_def, str):
                    slide_def = {"title": slide_def, "bullets": []}
                title   = slide_def.get("title",   f"Slide {i+1}") if isinstance(slide_def, dict) else f"Slide {i+1}"
                bullets = slide_def.get("bullets", [])              if isinstance(slide_def, dict) else []
                bullets = [str(b) for b in bullets if b]

                # Enrich thin slides
                if i > 0 and len(bullets) < 3:
                    try:
                        bullets = await client.generate_content(title)
                        log.ok(f"HF Content [{i}]", f"{len(bullets)} bullets")
                    except Exception as e:
                        log.error(f"HF Content [{i}]", str(e))
                        bullets = bullets or [
                            "Key concept overview",
                            "Important details and examples",
                            "Real-world applications",
                        ]

                # Image slide or normal slide
                if i == img_slide_index:
                    try:
                        emit(f"  🖼️ Slide {i+1}/{len(plan)}: *{title}* — generating HF image…")
                        res = await client.add_image_slide(
                            title=title,
                            image_prompt=f"High quality illustration of {title}, educational, detailed",
                            caption=f"AI-generated image: {title}",
                        )
                        slides_added += 1
                        emit(f"  ✅ {res}")
                        log.ok(f"ImageSlide {i+1}", res)
                    except Exception as e:
                        log.error(f"ImageSlide {i+1}", str(e))
                        emit(f"  ⚠️ Image failed — adding as content slide")
                        try:
                            res = await client.add_slide(title, bullets)
                            slides_added += 1
                            log.ok(f"Slide {i+1} (fallback)", res)
                        except Exception as e2:
                            log.error(f"Slide {i+1} fallback", str(e2))
                else:
                    try:
                        res = await client.add_slide(title, bullets)
                        slides_added += 1
                        emit(f"  ✅ Slide {i+1}/{len(plan)}: *{title}*")
                        log.ok(f"Slide {i+1}", res)
                    except Exception as e:
                        log.error(f"Slide {i+1}", str(e))
                        emit(f"  ⚠️ Slide {i+1} failed: {e}")

            # ── Step 4: Save ───────────────────────────────────────────────────
            emit("💾 Step 4 – Saving presentation…")
            save_res = await client.save_presentation(filename)
            log.ok("Save", save_res)
            emit(f"✅ {save_res}")

        # Context manager has exited cleanly here
        full_path = os.path.join(output_dir, filename)
        return {
            "success":  True,
            "filename": filename,
            "path":     full_path,
            "slides":   slides_added,
            "plan":     plan,
            "log":      log,
        }

    except Exception as e:
        log.error("Agent", str(e))
        return {"success": False, "error": str(e), "log": log}