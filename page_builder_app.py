import streamlit as st
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import json
import cssutils
import logging
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from io import BytesIO

# Disable CSS parsing warnings
cssutils.log.setLevel(logging.CRITICAL)
load_dotenv()

st.set_page_config(page_title="Edstellar Page Builder", page_icon="📝", layout="wide")

LIBRARY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Library")
LIBRARY_INDEX = os.path.join(LIBRARY_DIR, "library-index.json")

OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

DEEPSEEK_MODELS = {
    "DeepSeek V3": "deepseek/deepseek-chat-v3-0324",
    "DeepSeek R1": "deepseek/deepseek-r1",
    "DeepSeek V3 (free)": "deepseek/deepseek-chat-v3-0324:free",
    "DeepSeek R1 (free)": "deepseek/deepseek-r1:free",
}

# ── Helpers ──────────────────────────────────────────────────────────────────

def extract_docx_text(file_bytes):
    """Extract raw text from a .docx byte stream."""
    text = []
    with zipfile.ZipFile(file_bytes) as docx:
        xml_content = docx.read("word/document.xml")
        tree = ET.XML(xml_content)
        for node in tree.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
            if node.text:
                text.append(node.text)
    return "\n".join(text)


def parse_page_flow(docx_text):
    """Deterministically extract the ordered module list from the PAGE FLOW OVERVIEW table."""
    sections = []
    lines = docx_text.split("\n")

    start_idx = None
    end_idx = None
    for i, line in enumerate(lines):
        if "PAGE FLOW OVERVIEW" in line:
            start_idx = i
        if start_idx and "NOTES" == line.strip() and i > start_idx + 5:
            end_idx = i
            break
        if start_idx and line.strip().startswith("S01") or (start_idx and "________________" in line and i > start_idx + 10):
            if end_idx is None:
                end_idx = i
            break

    if start_idx is None:
        return []

    flow_text = "\n".join(lines[start_idx:end_idx] if end_idx else lines[start_idx:start_idx + 200])

    html_pattern = re.compile(r"(edstellar-[\w-]+\.html|seo-metadata-block-template\.md)")
    matches = html_pattern.findall(flow_text)

    number_pattern = re.compile(r"^(\d{2})$", re.MULTILINE)
    numbers = number_pattern.findall(flow_text)

    for idx, match in enumerate(matches):
        num = numbers[idx] if idx < len(numbers) else f"{idx + 1:02d}"
        sections.append({
            "number": num,
            "module_file": match,
            "is_html": match.endswith(".html"),
        })

    return sections


def parse_section_content(docx_text):
    """Parse the per-section content blocks from the docx."""
    sections = {}
    parts = re.split(r"_{10,}", docx_text)
    for part in parts:
        match = re.match(r"\s*S(\d{2})\s*\|\s*\n(.+?)(?:\n|$)", part)
        if match:
            sec_num = f"S{match.group(1)}"
            sections[sec_num] = part.strip()
    return sections


def scope_selector_list(selector_list, scope_id):
    """Prefix CSS selectors with a scope wrapper ID."""
    for selector in selector_list:
        text = selector.selectorText
        if text.startswith(":root") or text.startswith("body") or text.startswith("html"):
            selector.selectorText = (
                text.replace(":root", f"#{scope_id}")
                .replace("body", f"#{scope_id}")
                .replace("html", f"#{scope_id}")
            )
        else:
            selector.selectorText = f"#{scope_id} {text}"


def scope_module_css(html_content, scope_id):
    """Scope <style> tags and wrap body content in a scoped <div>."""
    soup = BeautifulSoup(html_content, "html.parser")
    styles = []
    for style in soup.find_all("style"):
        if not style.string:
            continue
        sheet = cssutils.parseString(style.string)
        for rule in sheet:
            if rule.type == rule.STYLE_RULE:
                scope_selector_list(rule.selectorList, scope_id)
            elif rule.type == rule.MEDIA_RULE:
                for inner_rule in rule.cssRules:
                    if inner_rule.type == inner_rule.STYLE_RULE:
                        scope_selector_list(inner_rule.selectorList, scope_id)
        styles.append(sheet.cssText.decode("utf-8"))
        style.decompose()

    body = soup.find("body")
    if body:
        inner_html = "".join(str(c) for c in body.contents)
    else:
        inner_html = str(soup)

    final_css = "\n".join(styles)
    return f'<style>\n{final_css}\n</style>\n<div id="{scope_id}">\n{inner_html}\n</div>'


def call_openrouter(api_key, model, prompt):
    """Call OpenRouter API with the given model and prompt. Returns the text response."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://edstellar.com",
        "X-Title": "Edstellar Page Builder",
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "You are an expert HTML developer. You only output valid HTML. No markdown fences, no commentary, no explanations."
            },
            {
                "role": "user",
                "content": prompt,
            }
        ],
        "max_tokens": 16000,
        "temperature": 0.2,
    }
    response = requests.post(OPENROUTER_API_URL, headers=headers, json=payload, timeout=120)

    # Parse response body before raising so we get the actual error message
    try:
        data = response.json()
    except Exception:
        response.raise_for_status()
        raise RuntimeError(f"Non-JSON response ({response.status_code}): {response.text[:500]}")

    if not response.ok:
        err_msg = data.get("error", {}).get("message", "") if isinstance(data.get("error"), dict) else str(data.get("error", ""))
        raise RuntimeError(f"OpenRouter {response.status_code}: {err_msg or response.text[:300]}")

    if "error" in data:
        raise RuntimeError(data["error"].get("message", str(data["error"])))

    text = data["choices"][0]["message"]["content"].strip()

    # DeepSeek R1 may include <think>...</think> blocks — strip them
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    # Strip markdown fences if present
    if text.startswith("```html"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()


def inject_content_with_deepseek(api_key, model, module_filename, scoped_html, section_content, section_number, full_docx_text):
    """Use DeepSeek via OpenRouter to replace placeholder text with real content."""
    prompt = f"""You are building an Edstellar consulting landing page.

Below is:
1. The FULL document content for reference.
2. The SPECIFIC section content block (Section {section_number}) from the developer reference document.
3. The HTML boilerplate for module `{module_filename}`.

Your job: Replace ALL placeholder/boilerplate text in the HTML with the actual content from the section content block. Map headings, paragraphs, list items, stats, CTAs, FAQ questions/answers, testimonials, etc. to their corresponding HTML elements.

RULES:
1. DO NOT change any CSS classes, inline styles, IDs, structural divs, or layout.
2. DO NOT add or remove HTML elements — only replace text content.
3. REPLACE placeholders like "Generic Headline", "Placeholder paragraph", "Lorem ipsum" etc. with content from the document.
4. For stats/numbers, ensure the exact values from the document are used.
5. Return ONLY the final valid HTML block. No markdown fences, no commentary.

── SECTION CONTENT BLOCK ({section_number}) ──
{section_content}

── FULL DOCUMENT (for cross-reference) ──
{full_docx_text[:12000]}

── HTML BOILERPLATE TO INJECT ──
{scoped_html}
"""
    return call_openrouter(api_key, model, prompt)


def load_library_index():
    """Load the library-index.json for section metadata."""
    if os.path.exists(LIBRARY_INDEX):
        with open(LIBRARY_INDEX, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def get_available_modules():
    """List all .html files in the Library folder."""
    if not os.path.exists(LIBRARY_DIR):
        return []
    return sorted(f for f in os.listdir(LIBRARY_DIR) if f.endswith(".html"))


# ── Streamlit UI ─────────────────────────────────────────────────────────────

st.title("📝 Edstellar Page Builder")
st.caption("Upload a Developer Reference .docx → build a fully stitched HTML page from the design library.")

# Sidebar config
with st.sidebar:
    st.header("⚙️ Configuration")

    api_key = os.getenv("OPENROUTER_API_KEY", "")
    if not api_key:
        api_key = st.text_input("OpenRouter API Key", type="password", help="Get your key at openrouter.ai/keys")
    else:
        st.success("OpenRouter API Key loaded from .env")

    selected_model_label = st.selectbox("DeepSeek Model", list(DEEPSEEK_MODELS.keys()), index=0)
    selected_model = DEEPSEEK_MODELS[selected_model_label]
    st.caption(f"`{selected_model}`")

    if api_key:
        if st.button("🔑 Test API Key"):
            try:
                test_headers = {
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                    "HTTP-Referer": "https://edstellar.com",
                    "X-Title": "Edstellar Page Builder",
                }
                test_payload = {
                    "model": selected_model,
                    "messages": [{"role": "user", "content": "Reply with just: OK"}],
                    "max_tokens": 5,
                }
                r = requests.post(OPENROUTER_API_URL, headers=test_headers, json=test_payload, timeout=30)
                rdata = r.json()
                if r.ok and "choices" in rdata:
                    st.success(f"Key valid! Model: {selected_model_label}")
                else:
                    err = rdata.get("error", {})
                    msg = err.get("message", str(err)) if isinstance(err, dict) else str(err)
                    st.error(f"API error: {msg}")
            except Exception as e:
                st.error(f"Connection failed: {e}")

    st.divider()
    st.header("📚 Design Library")
    lib_index = load_library_index()
    available = get_available_modules()
    st.metric("Available Modules", len(available))

    with st.expander("Browse Library Sections"):
        if lib_index and "sections" in lib_index:
            for key, sec in lib_index["sections"].items():
                is_visual = sec.get("is_visual", True)
                if not is_visual:
                    continue
                st.markdown(f"**{sec['label']}**")
                for v in sec.get("variants", []):
                    st.markdown(f"- `{v['file']}` — {v['layout'][:60]}")

# ── Main area ────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("Upload Developer Reference (.docx)", type="docx")

if uploaded_file and not api_key:
    st.warning("Enter your OpenRouter API Key in the sidebar to continue.")
    st.stop()

if not uploaded_file:
    st.info("Upload a `.docx` developer reference document to get started.")
    st.stop()

# Step 1: Extract text
if "docx_text" not in st.session_state or st.session_state.get("docx_filename") != uploaded_file.name:
    with st.spinner("Extracting text from document..."):
        st.session_state.docx_text = extract_docx_text(BytesIO(uploaded_file.read()))
        st.session_state.docx_filename = uploaded_file.name
        st.session_state.pop("page_flow", None)
        st.session_state.pop("final_html", None)

docx_text = st.session_state.docx_text

# Step 2: Parse page flow
if "page_flow" not in st.session_state:
    st.session_state.page_flow = parse_page_flow(docx_text)

page_flow = st.session_state.page_flow

if not page_flow:
    st.error("Could not find a PAGE FLOW OVERVIEW table in the document. Check document format.")
    st.stop()

# Step 3: Section editor
st.subheader("Page Flow")
st.caption(f"Found **{len(page_flow)}** sections in the document. Toggle sections on/off or swap modules.")

section_content_map = parse_section_content(docx_text)

cols_header = st.columns([0.5, 3, 4, 1])
cols_header[0].markdown("**#**")
cols_header[1].markdown("**Module File**")
cols_header[2].markdown("**Content Preview**")
cols_header[3].markdown("**Include**")

enabled_sections = []
for i, sec in enumerate(page_flow):
    if not sec["is_html"]:
        continue

    sec_key = f"S{sec['number']}"
    content_preview = ""
    if sec_key in section_content_map:
        content_lines = section_content_map[sec_key].split("\n")
        for cl in content_lines[2:6]:
            if cl.strip() and "DESIGN MODULE" not in cl and "edstellar-" not in cl:
                content_preview = cl.strip()[:80]
                break

    col = st.columns([0.5, 3, 4, 1])
    col[0].markdown(f"`{sec['number']}`")
    col[1].code(sec["module_file"], language=None)
    col[2].markdown(f"_{content_preview}_" if content_preview else "_—_")
    include = col[3].checkbox("", value=True, key=f"include_{i}", label_visibility="collapsed")

    if include:
        enabled_sections.append((i, sec))

st.divider()
st.markdown(f"**{len(enabled_sections)}** sections enabled for build.")

# Step 4: Build
if st.button("🚀 Build HTML Page", type="primary", use_container_width=True):
    if not api_key:
        st.error("No OpenRouter API key configured.")
        st.stop()

    progress = st.progress(0, text="Starting build...")
    status_area = st.empty()
    log_area = st.expander("Build Log", expanded=True)

    final_blocks = []
    total = len(enabled_sections)

    for step, (orig_idx, sec) in enumerate(enabled_sections):
        module_file = sec["module_file"]
        sec_num = f"S{sec['number']}"
        file_path = os.path.join(LIBRARY_DIR, module_file)

        progress.progress((step) / total, text=f"Processing {sec_num}: {module_file}")

        if not os.path.exists(file_path):
            log_area.warning(f"⚠️ {module_file} not found in Library — skipping.")
            continue

        with open(file_path, "r", encoding="utf-8") as f:
            raw_html = f.read()

        # Scope CSS
        scope_id = f"s{orig_idx:02d}_module"
        scoped_html = scope_module_css(raw_html, scope_id)
        log_area.info(f"✅ Scoped CSS for {module_file} → #{scope_id}")

        # Inject content via DeepSeek
        section_content = section_content_map.get(sec_num, "")
        if section_content:
            try:
                injected = inject_content_with_deepseek(
                    api_key, selected_model,
                    module_file, scoped_html, section_content, sec_num, docx_text
                )
                final_blocks.append(f"<!-- {sec_num}: {module_file} -->\n{injected}")
                log_area.success(f"✅ Content injected for {sec_num}: {module_file}")
            except Exception as e:
                final_blocks.append(f"<!-- {sec_num}: {module_file} (fallback) -->\n{scoped_html}")
                log_area.error(f"❌ DeepSeek injection failed for {sec_num}: {e} — using boilerplate.")
        else:
            final_blocks.append(f"<!-- {sec_num}: {module_file} -->\n{scoped_html}")
            log_area.warning(f"⚠️ No content block found for {sec_num} — using boilerplate.")

    progress.progress(1.0, text="Build complete!")

    # Assemble final HTML
    final_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Generated Consulting Page | Edstellar</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800&family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
</head>
<body>
{chr(10).join(final_blocks)}
</body>
</html>"""

    st.session_state.final_html = final_html
    status_area.success(f"Page built with {len(final_blocks)} sections using {selected_model_label}.")

# Step 5: Preview & Download
if "final_html" in st.session_state:
    st.divider()
    st.subheader("Output")

    tab_preview, tab_download, tab_source = st.tabs(["🔍 Preview", "📥 Download", "📄 Source"])

    with tab_preview:
        st.components.v1.html(st.session_state.final_html, height=800, scrolling=True)

    with tab_download:
        st.download_button(
            label="Download HTML File",
            data=st.session_state.final_html,
            file_name="generated-consulting-page.html",
            mime="text/html",
            use_container_width=True,
        )
        st.metric("File Size", f"{len(st.session_state.final_html) / 1024:.1f} KB")
        st.metric("Sections", len([b for b in st.session_state.final_html.split("<!--") if "module" in b.lower()]))

    with tab_source:
        st.code(st.session_state.final_html[:5000] + "\n\n... (truncated)", language="html")
