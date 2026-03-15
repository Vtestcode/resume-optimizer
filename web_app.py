import json
import os
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from app import (
    AIEngine,
    DocumentParser,
    DocumentWriter,
    ParagraphRebuilder,
    _normalize_line_text,
)


def inject_custom_theme():
    st.markdown(
        """
<style>
:root {
    --bg: #f6efe7;
    --paper: #fff8f0;
    --card: rgba(255, 249, 241, 0.88);
    --ink: #1f2937;
    --muted: #667085;
    --line: rgba(31, 41, 55, 0.09);
    --accent: #c56a3d;
    --accent-2: #235347;
    --accent-3: #134e5e;
    --shadow: 0 24px 60px rgba(59, 36, 20, 0.12);
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(197, 106, 61, 0.16), transparent 24%),
        radial-gradient(circle at top right, rgba(35, 83, 71, 0.14), transparent 22%),
        linear-gradient(180deg, #fdf8f2 0%, #f6efe7 100%);
    color: var(--ink);
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(18, 31, 46, 0.98), rgba(19, 78, 94, 0.95));
    border-right: 1px solid rgba(255, 255, 255, 0.08);
}

[data-testid="stSidebar"] * {
    color: #f7f3ef !important;
}

.block-container {
    max-width: 1220px;
    padding-top: 2.2rem;
    padding-bottom: 4rem;
}

.section-card {
    padding: 1.05rem 1.1rem 1.15rem;
    border-radius: 24px;
    border: 1px solid var(--line);
    background: var(--card);
    box-shadow: 0 16px 44px rgba(31, 41, 55, 0.08);
    backdrop-filter: blur(12px);
}

.section-kicker {
    font-size: 0.75rem;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--accent-3);
    font-weight: 800;
    margin-bottom: 0.55rem;
}

.section-title {
    font-size: 1.45rem;
    line-height: 1.12;
    font-weight: 800;
    color: var(--ink);
    margin-bottom: 0.4rem;
}

.section-copy {
    color: var(--muted);
    line-height: 1.6;
}

.pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 0.5rem;
    margin-top: 0.85rem;
}

.pill {
    padding: 0.42rem 0.78rem;
    border-radius: 999px;
    border: 1px solid rgba(31, 41, 55, 0.08);
    background: rgba(255, 255, 255, 0.82);
    font-size: 0.92rem;
}

.micro-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 0.8rem;
    margin-top: 0.85rem;
}

.micro-card {
    border-radius: 18px;
    border: 1px solid var(--line);
    background: rgba(255, 255, 255, 0.76);
    padding: 0.9rem;
}

.micro-label {
    display: block;
    color: var(--muted);
    font-size: 0.75rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
}

.micro-value {
    display: block;
    margin-top: 0.25rem;
    font-size: 1.35rem;
    font-weight: 800;
    color: var(--ink);
}

.result-band {
    margin: 1rem 0 1.2rem;
    padding: 1rem 1.15rem;
    border-radius: 24px;
    background: linear-gradient(135deg, rgba(19, 78, 94, 0.96), rgba(35, 83, 71, 0.94));
    color: white;
    box-shadow: var(--shadow);
}

.result-band h3,
.result-band p {
    margin: 0;
}

.result-band p {
    margin-top: 0.25rem;
    color: rgba(255, 255, 255, 0.82);
}

.diff-card {
    padding: 0.95rem 1rem;
    border-radius: 18px;
    border: 1px solid var(--line);
    background: rgba(255, 255, 255, 0.78);
    margin-bottom: 0.85rem;
}

.diff-old {
    color: #9b2226;
    text-decoration: line-through;
    margin-bottom: 0.45rem;
}

.diff-new {
    color: #14532d;
}

.diff-same {
    color: #475569;
}

div[data-testid="stFileUploader"],
div[data-testid="stTextArea"],
div[data-testid="stTextInput"],
div[data-testid="stSelectbox"],
div[data-testid="stMetric"],
div[data-testid="stDataFrame"] {
    background: rgba(255, 255, 255, 0.6);
    border-radius: 18px;
}

div[data-testid="stTextArea"] textarea {
    min-height: 360px;
}

@media (max-width: 900px) {
    .micro-grid {
        grid-template-columns: 1fr;
    }
}
</style>
        """,
        unsafe_allow_html=True,
    )


def render_section_header(kicker: str, title: str, copy: str):
    st.markdown(
        f"""
<div class="section-card">
  <div class="section-kicker">{kicker}</div>
  <div class="section-title">{title}</div>
  <div class="section-copy">{copy}</div>
</div>
        """,
        unsafe_allow_html=True,
    )


def render_signal_hero(api_ready: bool, uploaded_name: str, jd_text: str, has_results: bool):
    payload = {
        "apiReady": api_ready,
        "resumeName": uploaded_name or "No resume uploaded yet",
        "jdWords": len(jd_text.split()),
        "hasResults": has_results,
    }
    components.html(
        f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
body {{
    margin: 0;
    font-family: Arial, sans-serif;
    background: transparent;
}}
.hero {{
    position: relative;
    overflow: hidden;
    border-radius: 30px;
    padding: 30px;
    background:
        radial-gradient(circle at 18% 18%, rgba(197, 106, 61, 0.26), transparent 24%),
        radial-gradient(circle at 84% 24%, rgba(35, 83, 71, 0.22), transparent 22%),
        linear-gradient(135deg, rgba(255,249,241,0.98), rgba(247,240,230,0.96));
    border: 1px solid rgba(31, 41, 55, 0.08);
    box-shadow: 0 18px 48px rgba(31, 41, 55, 0.08);
}}
.eyebrow {{
    font-size: 12px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #134e5e;
    font-weight: 800;
}}
.title {{
    margin: 14px 0 10px;
    color: #1f2937;
    font-size: 44px;
    line-height: 1.02;
    font-weight: 800;
}}
.rotator {{
    color: #c56a3d;
}}
.copy {{
    max-width: 760px;
    color: #5d6675;
    line-height: 1.65;
    font-size: 16px;
}}
.stats {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
    margin-top: 18px;
}}
.stat {{
    border-radius: 18px;
    border: 1px solid rgba(31, 41, 55, 0.08);
    background: rgba(255,255,255,0.82);
    padding: 14px;
}}
.label {{
    display: block;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #667085;
}}
.value {{
    display: block;
    margin-top: 8px;
    font-size: 26px;
    font-weight: 800;
    color: #1f2937;
}}
.resume {{
    margin-top: 16px;
    padding: 13px 14px;
    border-radius: 16px;
    background: rgba(31, 41, 55, 0.04);
    color: #48576a;
    font-size: 14px;
}}
@media (max-width: 760px) {{
    .title {{ font-size: 31px; }}
    .stats {{ grid-template-columns: 1fr 1fr; }}
}}
</style>
</head>
<body>
<div class="hero">
  <div class="eyebrow">Resume optimizer studio</div>
  <div class="title">Tune every bullet for <span id="rotator" class="rotator"></span></div>
  <div class="copy">
    This front end keeps the Word layout intact, pushes the rewrite harder toward the pasted JD,
    and shows the result in a cleaner deployment-ready workflow.
  </div>
  <div class="stats">
    <div class="stat"><span class="label">API status</span><span class="value" id="apiMetric"></span></div>
    <div class="stat"><span class="label">JD words</span><span class="value" id="jdMetric"></span></div>
    <div class="stat"><span class="label">Pipeline mode</span><span class="value" id="flowMetric"></span></div>
    <div class="stat"><span class="label">Result state</span><span class="value" id="resultMetric"></span></div>
  </div>
  <div id="resumeName" class="resume"></div>
</div>
<script>
const payload = {json.dumps(payload)};
const words = ["grounded alignment", "quantified impact", "ATS signal", "high-match language"];
let idx = 0;
const rotator = document.getElementById("rotator");
const paintWord = () => {{
  rotator.textContent = words[idx % words.length];
  idx += 1;
}};
paintWord();
setInterval(paintWord, 1800);
document.getElementById("apiMetric").textContent = payload.apiReady ? "Ready" : "Missing";
document.getElementById("jdMetric").textContent = payload.jdWords ? payload.jdWords + "w" : "0w";
document.getElementById("flowMetric").textContent = payload.jdWords > 250 ? "Deep" : "Draft";
document.getElementById("resultMetric").textContent = payload.hasResults ? "Loaded" : "Pending";
document.getElementById("resumeName").textContent = "Resume file: " + payload.resumeName;
</script>
</body>
</html>
        """,
        height=320,
    )


def render_results_dashboard(result: Dict[str, Any]):
    ats = result.get("ats", {})
    payload = {
        "score": int(ats.get("ats_score", 0)),
        "missing": len(ats.get("missing_skills", [])),
        "keywords": len(ats.get("keyword_freq", {})),
        "summary": result.get("summary_count", 0),
        "experience": result.get("experience_count", 0),
    }
    components.html(
        f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
body {{ margin: 0; font-family: Arial, sans-serif; background: transparent; }}
.board {{
    display: grid;
    grid-template-columns: 220px 1fr;
    gap: 14px;
}}
.gauge, .panel {{
    border-radius: 24px;
    padding: 20px;
    background: linear-gradient(180deg, rgba(255,249,241,0.98), rgba(247,240,230,0.95));
    border: 1px solid rgba(31, 41, 55, 0.08);
}}
.ring-wrap {{
    position: relative;
    width: 150px;
    height: 150px;
    margin: 0 auto;
}}
.ring {{
    width: 150px;
    height: 150px;
    border-radius: 999px;
    background: conic-gradient(#235347 calc(var(--score) * 1%), rgba(19, 78, 94, 0.12) 0);
}}
.ring-hole {{
    position: absolute;
    inset: 19px;
    display: grid;
    place-items: center;
    border-radius: 999px;
    background: #fff9f1;
}}
.score {{
    text-align: center;
    color: #1f2937;
    font-size: 32px;
    font-weight: 800;
}}
.score small {{
    display: block;
    margin-top: 4px;
    font-size: 12px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #667085;
}}
.grid {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
}}
.tile {{
    border-radius: 18px;
    padding: 14px;
    background: rgba(255,255,255,0.78);
    border: 1px solid rgba(31, 41, 55, 0.08);
}}
.tile strong {{
    display: block;
    font-size: 28px;
    color: #1f2937;
}}
.tile span {{
    display: block;
    margin-top: 6px;
    color: #667085;
    font-size: 12px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
}}
@media (max-width: 760px) {{
    .board {{ grid-template-columns: 1fr; }}
    .grid {{ grid-template-columns: 1fr 1fr; }}
}}
</style>
</head>
<body>
<div class="board">
  <div class="gauge" style="--score: {payload["score"]};">
    <div class="ring-wrap">
      <div class="ring"></div>
      <div class="ring-hole">
        <div class="score"><span id="scoreValue">0</span><small>ATS score</small></div>
      </div>
    </div>
  </div>
  <div class="panel">
    <div class="grid">
      <div class="tile"><strong>{payload["experience"]}</strong><span>Experience bullets</span></div>
      <div class="tile"><strong>{payload["summary"]}</strong><span>Summary bullets</span></div>
      <div class="tile"><strong>{payload["keywords"]}</strong><span>Tracked keywords</span></div>
      <div class="tile"><strong>{payload["missing"]}</strong><span>Missing skills</span></div>
    </div>
  </div>
</div>
<script>
const target = {payload["score"]};
const el = document.getElementById("scoreValue");
let current = 0;
const step = Math.max(1, Math.ceil(target / 30));
const timer = setInterval(() => {{
  current = Math.min(target, current + step);
  el.textContent = current;
  if (current >= target) clearInterval(timer);
}}, 25);
</script>
</body>
</html>
        """,
        height=240,
    )


def render_diff_blocks(result: Dict[str, Any]):
    if result["summary_bullets"]:
        st.markdown("### Professional Summary")
        for binfo in result["summary_bullets"]:
            key = (binfo["para_idx"], binfo["line_idx"])
            orig = result["original_map"].get(key, binfo["text"])
            fallback_text, is_bullet = result["all_updates"].get(key, (orig, True))
            shown_text = result["applied_map"].get(key, fallback_text)
            clean_new = _normalize_line_text(shown_text, is_bullet)
            if _normalize_line_text(orig, True) != clean_new:
                st.markdown(
                    f'<div class="diff-card"><div class="diff-old">{orig[:200]}</div><div class="diff-new">{shown_text[:200]}</div></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="diff-card"><div class="diff-same">{orig[:200]} (unchanged)</div></div>',
                    unsafe_allow_html=True,
                )

    for rk, rv in result["roles"].items():
        st.markdown(f"### {rk[:120]}")
        for binfo in rv["bullets"]:
            key = (binfo["para_idx"], binfo["line_idx"])
            orig = result["original_map"].get(key, binfo["text"])
            fallback_text, is_bullet = result["all_updates"].get(key, (orig, True))
            shown_text = result["applied_map"].get(key, fallback_text)
            clean_new = _normalize_line_text(shown_text, is_bullet)
            if _normalize_line_text(orig, True) != clean_new:
                st.markdown(
                    f'<div class="diff-card"><div class="diff-old">{orig[:200]}</div><div class="diff-new">{shown_text[:200]}</div></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="diff-card"><div class="diff-same">{orig[:200]} (unchanged)</div></div>',
                    unsafe_allow_html=True,
                )


def run_alignment(api_key: str, model: str, uploaded, jd: str) -> Dict[str, Any]:
    file_bytes = uploaded.read()
    ai = AIEngine(api_key=api_key, model=model)

    with st.status("Parsing resume structure...", expanded=True) as status:
        doc = DocumentParser.load_document(file_bytes)
        parsed = DocumentParser.parse_paragraphs(doc)
        roles = DocumentParser.extract_roles(parsed)
        summary_bullets = DocumentParser.extract_summary_bullets(parsed)
        skills = DocumentParser.extract_skills(parsed)
        header_points = DocumentParser.extract_header_points(parsed)
        total_b = sum(len(role["bullets"]) for role in roles.values())
        st.write(
            f"Detected {len(roles)} roles, {total_b} experience bullets, "
            f"{len(summary_bullets)} summary bullets, and {len(skills)} skill lines."
        )
        status.update(label="Resume map captured", state="complete")

    all_updates: Dict[Tuple[int, int], Tuple[str, bool]] = {}
    original_map: Dict[Tuple[int, int], str] = {}

    with st.status("Rewriting high-visibility content...", expanded=True) as status:
        if header_points:
            for hinfo in header_points:
                key = (hinfo["para_idx"], hinfo["line_idx"])
                try:
                    rewritten_header = ai.rewrite_header_points(hinfo["text"], jd)
                except Exception as e:
                    st.warning(f"Header strip error: {e}")
                    rewritten_header = hinfo["text"]
                all_updates[key] = (rewritten_header, False)
                original_map[key] = hinfo["text"]
        if summary_bullets:
            originals = [bullet["text"] for bullet in summary_bullets]
            try:
                rewritten = ai.rewrite_summary(originals, jd, aggressive_rewrite=True)
            except Exception as e:
                st.warning(f"Summary error: {e}")
                rewritten = originals
            for binfo, new_t in zip(summary_bullets, rewritten):
                key = (binfo["para_idx"], binfo["line_idx"])
                all_updates[key] = (new_t, True)
                original_map[key] = binfo["text"]
        status.update(label="Header and summary aligned", state="complete")

    with st.status("Optimizing experience bullets...", expanded=True) as status:
        progress = st.progress(0.0)
        role_items = list(roles.items())
        for index, (role_key, role_value) in enumerate(role_items):
            if not role_value["bullets"]:
                progress.progress((index + 1) / max(len(role_items), 1))
                continue
            st.write(f"Processing {role_key[:80]}")
            originals = [bullet["text"] for bullet in role_value["bullets"]]
            aggressive_rewrite = index < 2
            try:
                rewritten = ai.rewrite_bullets(
                    role_key,
                    originals,
                    jd,
                    aggressive_rewrite=aggressive_rewrite,
                )
            except Exception as e:
                st.warning(f"Error for '{role_key[:40]}': {e}")
                rewritten = originals
            for binfo, new_t in zip(role_value["bullets"], rewritten):
                key = (binfo["para_idx"], binfo["line_idx"])
                all_updates[key] = (new_t, True)
                original_map[key] = binfo["text"]
            progress.progress((index + 1) / max(len(role_items), 1))
        status.update(label=f"Aligned {total_b} bullets", state="complete")

    with st.status("Refreshing skill inventory...", expanded=True) as status:
        if skills:
            current_skill_lines = [skill["text"] for skill in skills]
            try:
                new_lines = ai.update_skills(current_skill_lines, jd)
                for skill_info, new_line in zip(skills, new_lines):
                    key = (skill_info["para_idx"], skill_info["line_idx"])
                    all_updates[key] = (new_line, False)
            except Exception as e:
                st.warning(f"Skills error: {e}")
        status.update(label="Skills refreshed", state="complete")

    with st.status("Rebuilding the Word document...", expanded=True) as status:
        original_doc = DocumentParser.load_document(file_bytes)
        aligned_doc = DocumentParser.load_document(file_bytes)
        updates_by_para: Dict[int, Dict[int, Tuple[str, bool]]] = {}
        for (para_idx, line_idx), (text, is_bullet) in all_updates.items():
            updates_by_para.setdefault(para_idx, {})[line_idx] = (text, is_bullet)

        for para_idx in sorted(updates_by_para.keys()):
            para = aligned_doc.paragraphs[para_idx]
            line_map = updates_by_para[para_idx]
            bullet_replacements = {line: text for line, (text, is_bullet) in line_map.items() if is_bullet}
            skill_replacements = {line: text for line, (text, is_bullet) in line_map.items() if not is_bullet}
            if bullet_replacements and skill_replacements:
                ParagraphRebuilder.rebuild(para, bullet_replacements, is_bullet=True)
                ParagraphRebuilder.rebuild(para, skill_replacements, is_bullet=False, is_skill=True)
            elif bullet_replacements:
                ParagraphRebuilder.rebuild(para, bullet_replacements, is_bullet=True)
            elif skill_replacements:
                ParagraphRebuilder.rebuild(para, skill_replacements, is_bullet=False, is_skill=True)

        DocumentWriter.sync_paragraph_layout(original_doc, aligned_doc)
        applied_map = DocumentWriter.verify_updates(aligned_doc, all_updates)
        doc_bytes = DocumentWriter.save_to_bytes(aligned_doc)
        status.update(label="Aligned document package ready", state="complete")

    with st.status("Scoring ATS and drafting cover letter...", expanded=True) as status:
        aligned_text = DocumentParser.full_text(DocumentParser.load_document(doc_bytes))
        try:
            ats = ai.ats_analysis(aligned_text, jd)
        except Exception:
            ats = {"ats_score": 0, "missing_skills": [], "keyword_freq": {}}
        try:
            cover = ai.generate_cover_letter(aligned_text, jd)
        except Exception as e:
            cover = f"Error: {e}"
        status.update(label="Analysis complete", state="complete")

    return {
        "doc_bytes": doc_bytes,
        "ats": ats,
        "cover": cover,
        "roles": roles,
        "summary_bullets": summary_bullets,
        "all_updates": all_updates,
        "original_map": original_map,
        "applied_map": applied_map,
        "summary_count": len(summary_bullets),
        "experience_count": total_b,
    }


def main():
    st.set_page_config(page_title="Resume Optimizer", page_icon="R", layout="wide")
    inject_custom_theme()

    if "result_bundle" not in st.session_state:
        st.session_state["result_bundle"] = None
    if "last_uploaded_name" not in st.session_state:
        st.session_state["last_uploaded_name"] = ""
    api_key = os.getenv("OPENAI_API_KEY", "")
    model = os.getenv("OPENAI_MODEL", "gpt-4o")

    uploaded_name = st.session_state.get("last_uploaded_name", "")
    jd_text = st.session_state.get("jd_input", "")
    render_signal_hero(
        api_ready=bool(api_key),
        uploaded_name=uploaded_name,
        jd_text=jd_text,
        has_results=st.session_state["result_bundle"] is not None,
    )

    left, right = st.columns([1.02, 1.18], gap="large")
    with left:
        render_section_header(
            "Command center",
            "Load the original resume",
            "Use the exact Word document so the rebuild can preserve formatting while the model optimizes the language underneath.",
        )
        uploaded = st.file_uploader("Resume (.docx)", type=["docx"], key="resume_upload")
        if uploaded:
            st.session_state["last_uploaded_name"] = uploaded.name
        st.markdown(
            """
<div class="pill-row">
  <div class="pill">Word formatting preserved</div>
  <div class="pill">Bullet grounding protected</div>
  <div class="pill">JD wording emphasized</div>
</div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        render_section_header(
            "JD workbench",
            "Paste the full target job description",
            "Long, detailed descriptions help the engine find required platforms, outcomes, and terminology before it starts rewriting.",
        )
        jd = st.text_area(
            "Job Description",
            key="jd_input",
            placeholder="Paste the full job description here...",
        )
        st.markdown(
            f"""
<div class="micro-grid">
  <div class="micro-card"><span class="micro-label">JD word count</span><span class="micro-value">{len(jd.split())}</span></div>
  <div class="micro-card"><span class="micro-label">Resume uploaded</span><span class="micro-value">{"Yes" if uploaded else "No"}</span></div>
  <div class="micro-card"><span class="micro-label">Run readiness</span><span class="micro-value">{"Ready" if api_key and uploaded else "Waiting"}</span></div>
</div>
            """,
            unsafe_allow_html=True,
        )

    action_left, action_mid, action_right = st.columns([1.15, 0.8, 0.8], gap="medium")
    with action_left:
        run_clicked = st.button("Align Resume", type="primary", use_container_width=True)
    with action_mid:
        if st.button("Clear Results", use_container_width=True):
            st.session_state["result_bundle"] = None
            st.rerun()
    with action_right:
        st.markdown(
            f"""
<div class="section-card">
  <div class="section-kicker">State</div>
  <div class="section-copy">{'Ready to run' if api_key and uploaded and jd.strip() else 'Waiting for required inputs'}</div>
</div>
            """,
            unsafe_allow_html=True,
        )

    if run_clicked:
        if not api_key:
            st.error("Add an OpenAI API key before running the optimizer.")
        elif not uploaded or not jd.strip():
            st.error("Upload a resume and paste the full job description before running.")
        else:
            try:
                st.session_state["result_bundle"] = run_alignment(api_key, model, uploaded, jd)
            except Exception as e:
                st.error(f"Alignment failed: {e}")

    result = st.session_state.get("result_bundle")
    if not result:
        st.markdown(
            """
<div class="result-band">
  <h3>Studio idle</h3>
  <p>No local process is active for this app right now. The workspace is clear and ready for Heroku deployment.</p>
</div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        """
<div class="result-band">
  <h3>Optimization complete</h3>
  <p>Your aligned resume, ATS readout, cover letter, and bullet-by-bullet diff are ready below.</p>
</div>
        """,
        unsafe_allow_html=True,
    )
    render_results_dashboard(result)

    overview_tab, ats_tab, cover_tab, diff_tab = st.tabs(
        ["Overview", "ATS & Keywords", "Cover Letter", "Bullet Diff"]
    )

    with overview_tab:
        st.download_button(
            "Download aligned resume (.docx)",
            data=result["doc_bytes"],
            file_name="aligned_resume.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
        st.markdown(
            f"""
<div class="micro-grid">
  <div class="micro-card"><span class="micro-label">Experience bullets</span><span class="micro-value">{result['experience_count']}</span></div>
  <div class="micro-card"><span class="micro-label">Summary bullets</span><span class="micro-value">{result['summary_count']}</span></div>
  <div class="micro-card"><span class="micro-label">ATS score</span><span class="micro-value">{result['ats'].get('ats_score', 0)}</span></div>
</div>
            """,
            unsafe_allow_html=True,
        )

    with ats_tab:
        score = result["ats"].get("ats_score", 0)
        st.metric("Match Score", f"{score} / 100")
        missing = result["ats"].get("missing_skills", [])
        if missing:
            st.warning("These JD skills are still missing from the resume:")
            for skill in missing:
                st.markdown(f"- {skill}")
        else:
            st.success("No critical missing skills detected.")

        keyword_freq = result["ats"].get("keyword_freq", {})
        if keyword_freq:
            df = pd.DataFrame([{"Keyword": key, "Count": value} for key, value in keyword_freq.items()])
            df = df.sort_values("Count", ascending=False).reset_index(drop=True)
            st.dataframe(df, use_container_width=True, hide_index=True)

    with cover_tab:
        st.text_area("Tailored cover letter", value=result["cover"], height=380)

    with diff_tab:
        render_diff_blocks(result)


if __name__ == "__main__":
    main()
