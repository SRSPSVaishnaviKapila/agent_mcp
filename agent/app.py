"""
Auto-PPT Agent – Streamlit UI
Run: streamlit run agent/app.py
"""

import asyncio
import os
import sys
import tempfile

import streamlit as st

# Ensure project root is on sys.path
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from agent.agent_core import run_agent

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auto-PPT Agent",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #065A82 0%, #02C39A 100%);
        padding: 2rem; border-radius: 12px;
        margin-bottom: 1.5rem; text-align: center; color: white;
    }
    .main-header h1 { margin: 0; font-size: 2.4rem; }
    .main-header p  { margin: 0.4rem 0 0; opacity: 0.85; font-size: 1rem; }
    .tool-badge {
        display: inline-block; background: #065A82; color: #02C39A;
        padding: 2px 10px; border-radius: 999px;
        font-size: 0.78rem; margin: 2px; font-family: monospace;
    }
    .plan-card {
        background: #f0fafa; border-left: 4px solid #02C39A;
        padding: 0.6rem 1rem; margin: 0.3rem 0;
        border-radius: 0 8px 8px 0;
    }
    .hf-badge {
        background: #FFD21E; color: #1A1A2E;
        padding: 3px 12px; border-radius: 999px;
        font-size: 0.8rem; font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🎯 Auto-PPT Agent</h1>
  <p>Powered by <span class="hf-badge">🤗 HuggingFace</span>
     — AI-generated content &amp; images, zero manual work</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    num_slides = st.slider("Number of slides", 3, 12, 5)

    include_image_slide = st.checkbox(
        "Include AI image slide",
        value=True,
        help="Adds one slide with a real HuggingFace-generated image"
    )

    st.divider()
    st.subheader("🤗 HuggingFace Config")

    hf_token = st.text_input(
        "HuggingFace API Token",
        type="password",
        value=os.getenv("HF_API_TOKEN", ""),
        placeholder="hf_...",
        help="Get free token at huggingface.co/settings/tokens"
    )
    if hf_token:
        os.environ["HF_API_TOKEN"] = hf_token

    hf_text_model = st.selectbox(
        "Text Model (content generation)",
        [
            "mistralai/Mistral-7B-Instruct-v0.3",
            "HuggingFaceH4/zephyr-7b-beta",
            "microsoft/Phi-3-mini-4k-instruct",
        ],
        help="Used for slide plans and bullet points"
    )
    os.environ["HF_TEXT_MODEL"] = hf_text_model

    hf_img_model = st.selectbox(
        "Image Model (slide images)",
        [
            "stabilityai/stable-diffusion-xl-base-1.0",
            "runwayml/stable-diffusion-v1-5",
            "prompthero/openjourney",
        ],
        help="Used to generate images for image slides"
    )
    os.environ["HF_IMG_MODEL"] = hf_img_model

    st.divider()
    st.subheader("🛠️ MCP Tools")
    for tool in ["create_presentation", "plan_slides", "generate_content",
                 "add_slide", "add_image_slide", "save_presentation"]:
        st.markdown(f'<span class="tool-badge">{tool}</span>', unsafe_allow_html=True)

    st.divider()
    st.caption("Auto-PPT Agent · AI Agents & MCP Architecture")

# ── Main input ─────────────────────────────────────────────────────────────────
col1, col2 = st.columns([4, 1])
with col1:
    user_prompt = st.text_input(
        "Describe your presentation",
        placeholder='e.g. "Create a 5-slide presentation on the solar system for 6th graders"',
        label_visibility="collapsed",
    )
with col2:
    generate_btn = st.button("✨ Generate", use_container_width=True, type="primary")

# ── Example prompts ────────────────────────────────────────────────────────────
with st.expander("💡 Example prompts"):
    examples = [
        "Create a 5-slide presentation on the life cycle of a star for 6th grade",
        "Make a 6-slide deck on the history of artificial intelligence",
        "Build a 4-slide presentation on climate change solutions",
        "Create a 5-slide pitch deck for a mobile app startup",
        "Make a presentation on photosynthesis for high school students",
    ]
    for ex in examples:
        if st.button(ex, key=ex):
            st.session_state["_example"] = ex
            st.rerun()

if "_example" in st.session_state and not user_prompt:
    user_prompt = st.session_state.pop("_example")

# ── Validate token before running ──────────────────────────────────────────────
if generate_btn and user_prompt.strip():
    if not os.getenv("HF_API_TOKEN", "").strip():
        st.error("❌ Please enter your HuggingFace API Token in the sidebar before generating.")
        st.stop()

    prompt = user_prompt.strip()
    if str(num_slides) not in prompt and "slide" not in prompt.lower():
        prompt = f"Create a {num_slides}-slide presentation on: {prompt}"

    # Pass image-slide setting to agent via env
    os.environ["INCLUDE_IMAGE_SLIDE"] = "1" if include_image_slide else "0"

    st.divider()
    st.subheader("🤖 Agent Progress")

    progress_placeholder = st.empty()
    log_lines: list[str] = []

    def on_progress(msg: str):
        log_lines.append(msg)
        progress_placeholder.markdown("\n\n".join(log_lines))

    with tempfile.TemporaryDirectory() as tmpdir:
        with st.spinner("🤗 HuggingFace is generating your presentation…"):
            # Use a dedicated event loop to avoid ExceptionGroup crashes
            # on Windows with Python 3.11+ and MCP stdio TaskGroups
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                result = loop.run_until_complete(
                    run_agent(prompt, output_dir=tmpdir, progress_cb=on_progress)
                )
            finally:
                loop.close()
                asyncio.set_event_loop(None)

        if result.get("success"):
            st.success(f"✅ Done! **{result['slides']}** slides created.")

            # Slide plan
            with st.expander("📋 Slide Plan", expanded=True):
                for i, s in enumerate(result.get("plan", [])):
                    label = "🎬 Title Slide" if i == 0 else f"Slide {i+1}"
                    st.markdown(
                        f'<div class="plan-card"><b>{label}:</b> {s["title"]}</div>',
                        unsafe_allow_html=True,
                    )

            # Download
            pptx_path = result["path"]
            if os.path.exists(pptx_path):
                with open(pptx_path, "rb") as f:
                    pptx_bytes = f.read()
                st.download_button(
                    label=f"⬇️ Download {result['filename']}",
                    data=pptx_bytes,
                    file_name=result["filename"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

            # Step log
            with st.expander("📜 Detailed Step Log"):
                step_log = result.get("log")
                if step_log:
                    for entry in step_log.entries:
                        icon = {"ok": "✅", "error": "❌", "info": "ℹ️"}.get(entry["status"], "•")
                        st.markdown(f"`{entry['time']}` {icon} **{entry['step']}** — {entry['detail']}")
        else:
            st.error(f"❌ Agent failed: {result.get('error', 'Unknown error')}")
            step_log = result.get("log")
            if step_log and step_log.entries:
                with st.expander("📜 Error Log"):
                    for entry in step_log.entries:
                        icon = {"ok": "✅", "error": "❌", "info": "ℹ️"}.get(entry["status"], "•")
                        st.markdown(f"`{entry['time']}` {icon} **{entry['step']}** — {entry['detail']}")

elif generate_btn:
    st.warning("⚠️ Please enter a presentation prompt first.")