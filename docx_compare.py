#!/usr/bin/env python3
"""
DOCX Compare Tool - Web Edition
A browser-based GUI for comparing two Word (.docx) files or arbitrary text.

Launches a local HTTP server and opens the browser.
Chinese text renders perfectly via the browser's native font support.

Features:
- Side-by-side paragraph-level file comparison
- Manual text comparison (type or paste text directly)
- Character-level inline diff highlighting with CJK normalization
- Table content extraction and comparison
- Synchronized scrolling
- Customizable colors (foreground + background) for diff highlighting
- Import/export color schemes as JSON
- Export comparison results

Usage:
    python docx_compare.py
"""

import difflib
import json
import os
import sys
import unicodedata
import webbrowser
import http.server
import socketserver
import threading
import urllib.parse
from datetime import datetime

try:
    from docx import Document
except ImportError:
    print("Please install python-docx: pip install python-docx")
    sys.exit(1)

try:
    import regex
except ImportError:
    print("Please install regex: pip install regex")
    sys.exit(1)


PORT = 18909
HTML_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_docx_compare_web.html")


# ──────────────────────── CJK Normalization ────────────────────────

_OVERRIDE_CHARS = {
    0x2F800: 0x4E3D, 0x2F804: 0x4FAE, 0x2F80D: 0x500B,
    0x2F80E: 0x500B, 0x2F813: 0x5141, 0x2F815: 0x5168,
    0x2F81A: 0x516D, 0x2F822: 0x5207, 0x2F82B: 0x52F9,
    0x2F82C: 0x5315, 0x2F835: 0x5374, 0x2F83B: 0x53E2,
    0x2F83F: 0x542F, 0x2F840: 0x5433, 0x2F848: 0x54B3,
    0x2F84B: 0x5510, 0x2F84C: 0x5511, 0x2F851: 0x56DB,
    0x2F852: 0x56E3, 0x2F860: 0x59C9, 0x2F863: 0x5A66,
    0x2F86B: 0x5B97, 0x2F874: 0x5C66, 0x2F877: 0x5C8C,
    0x2F884: 0x5E7A, 0x2F894: 0x5F8C, 0x2F8A0: 0x60AB,
    0x2F8A2: 0x610D, 0x2F8A8: 0x6144, 0x2F8B0: 0x6182,
    0x2F8B5: 0x61B2, 0x2F8B8: 0x61D4, 0x2F8C0: 0x6323,
    0x2F8F5: 0x6C88, 0x2F907: 0x728A, 0x2F93C: 0x767B,
    0x2F940: 0x76CA, 0x2F945: 0x76F4, 0x2F94E: 0x77A5,
    0x2F952: 0x77DB, 0x2F95B: 0x7950, 0x2F96B: 0x79E6,
    0x2F974: 0x7A33, 0x2F99D: 0x7BC0, 0x2F9A7: 0x7C3E,
    0x2F9BA: 0x7D2F, 0x2F9C7: 0x7D7A, 0x2F9D5: 0x7E9F,
    0x2F9DC: 0x7F4A, 0x2F9E7: 0x7FB9, 0x2FA0D: 0x819A,
    0x2FA18: 0x8340, 0x2FA1D: 0x83CC,
}


def normalize_cjk(text):
    """Normalize CJK text: NFKC + compatibility ideograph mapping."""
    nfkc = unicodedata.normalize("NFKC", text)
    chars = []
    for ch in nfkc:
        cp = ord(ch)
        mapped = _OVERRIDE_CHARS.get(cp, cp)
        chars.append(chr(mapped))
    return "".join(chars)


# ──────────────────────── Document Extractor ────────────────────────
class DocxExtractor:
    # ── Tracked Changes (修订) Namespace ──
    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    @staticmethod
    def _resolve_para_text(para_elem):
        """Extract paragraph text with track-changes resolved.

        Rules:
        1. <w:ins> child → accept the inserted text (include its <w:r> children)
        2. <w:del> child → skip entirely (exclude its <w:r> children)
        3. Regular <w:r> → include as-is
        4. Duplicate runs (same text repeated with different formatting) →
           kept by the caller's dedup logic, not here.
        """
        w = DocxExtractor.NS_W
        parts = []
        for child in para_elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "r":
                # Normal run — extract w:t text
                for node in child.iter():
                    ntag = node.tag.split("}")[-1] if "}" in node.tag else node.tag
                    if ntag == "t" and node.text:
                        parts.append(node.text)
                        break
            elif tag == "ins":
                # Tracked insertion — accept the inserted runs
                for r_elem in child.iter():
                    rtag = r_elem.tag.split("}")[-1] if "}" in r_elem.tag else r_elem.tag
                    if rtag == "t" and r_elem.text:
                        parts.append(r_elem.text)
                        break
            elif tag == "del":
                # Tracked deletion — skip entirely
                continue
            # Other elements (pPr, rPr, bookmarkStart, etc.) → ignore
        return "".join(parts)

    @staticmethod
    def extract_text(filepath):
        doc = Document(filepath)
        paragraphs = []
        body = doc.element.body
        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                text = DocxExtractor._resolve_para_text(child)
                if text:
                    paragraphs.append(text)
            elif tag == "tbl":
                table = None
                for t in doc.tables:
                    if t._element is child:
                        table = t
                        break
                if table:
                    for row in table.rows:
                        cells = [cell.text.strip() for cell in row.cells]
                        deduped = []
                        for c in cells:
                            if not deduped or c != deduped[-1]:
                                deduped.append(c)
                        row_text = " | ".join(deduped)
                        if row_text.replace("|", "").strip():
                            paragraphs.append(row_text)
        return paragraphs


# ──────────────────────── Diff Engine ────────────────────────
class DiffEngine:
    """Diff engine with CJK-aware text normalization.

    Matching logic:
    1. Empty paragraphs are skipped entirely.
    2. Before diffing, compares the first 5 CJK characters of each paragraph.
       If they don't match at all, the paragraph pair is treated as a full
       replace (no inline diff needed), and the algorithm moves to the next.
    """

    @staticmethod
    def _is_empty(text):
        """Check if text is empty or contains only whitespace/punctuation."""
        return not text or not text.strip()

    @staticmethod
    def _first_5_cjk(text):
        """Return first 5 CJK characters (assumes text is already normalized)."""
        cjk_chars = []
        for ch in text:
            if '\u4e00' <= ch <= '\u9fff' or '\u3400' <= ch <= '\u4dbf' or '\uf900' <= ch <= '\ufaff' or '\U00020000' <= ch <= '\U0002a6df' or '\U0002a700' <= ch <= '\U0002ebef' or '\U00030000' <= ch <= '\U0003134f':
                cjk_chars.append(ch)
                if len(cjk_chars) >= 5:
                    break
        return "".join(cjk_chars)

    @staticmethod
    def compare_paragraphs(left_paras, right_paras):
        # Filter out empty paragraphs
        left_filtered = [p for p in left_paras if not DiffEngine._is_empty(p)]
        right_filtered = [p for p in right_paras if not DiffEngine._is_empty(p)]

        # Normalize for matching
        left_norm = [normalize_cjk(p) for p in left_filtered]
        right_norm = [normalize_cjk(p) for p in right_filtered]

        # Let difflib do the alignment — it handles insert/delete/replace correctly
        sm = difflib.SequenceMatcher(None, left_norm, right_norm, autojunk=False)
        result = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for k in range(i2 - i1):
                    result.append(("equal", left_filtered[i1 + k], right_filtered[j1 + k]))
            elif tag == "replace":
                lg = left_filtered[i1:i2]
                rg = right_filtered[j1:j2]
                for k in range(max(len(lg), len(rg))):
                    lt = lg[k] if k < len(lg) else ""
                    rt = rg[k] if k < len(rg) else ""
                    result.append(("replace", lt, rt))
            elif tag == "delete":
                for k in range(i1, i2):
                    result.append(("remove", left_filtered[k], ""))
            elif tag == "insert":
                for k in range(j1, j2):
                    result.append(("add", "", right_filtered[k]))
        return result

    @staticmethod
    def inline_diff(left_text, right_text):
        if not left_text and not right_text:
            return []
        if not left_text or not right_text:
            return [("remove", left_text), ("add", right_text)]
        ln = normalize_cjk(left_text)
        rn = normalize_cjk(right_text)
        sm = difflib.SequenceMatcher(None, ln, rn, autojunk=False)
        segments = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                segments.append(("equal", left_text[i1:i2]))
            elif tag == "replace":
                if i1 < i2:
                    segments.append(("remove", left_text[i1:i2]))
                if j1 < j2:
                    segments.append(("add", right_text[j1:j2]))
            elif tag == "delete":
                segments.append(("remove", left_text[i1:i2]))
            elif tag == "insert":
                segments.append(("add", right_text[j1:j2]))
        return segments


# ──────────────────────── HTML Generator ────────────────────────
def generate_html():
    """Generate the full HTML application page."""
    return r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DOCX Compare Tool</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'SimHei', 'Noto Sans SC', 'PingFang SC', sans-serif; background: #f5f5f5; color: #333; height: 100vh; display: flex; flex-direction: column; font-size: 14px; }
.tab-bar { background: #e0e0e0; display: flex; padding: 0; border-bottom: 2px solid #4a90d9; }
.tab-btn { background: #d0d0d0; border: none; padding: 10px 24px; cursor: pointer; font-size: 14px; font-weight: 500; color: #444; transition: background 0.15s; }
.tab-btn:hover { background: #c0c0c0; }
.tab-btn.active { background: #fff; color: #000; border-bottom: 2px solid #4a90d9; margin-bottom: -2px; }
.tab-content { display: none; flex: 1; flex-direction: column; padding: 8px; overflow: hidden; }
.tab-content.active { display: flex; }

/* Toolbar */
.toolbar { display: flex; align-items: center; gap: 6px; padding: 8px 6px; background: #f0f0f0; flex-wrap: wrap; }
.toolbar label { font-weight: 500; white-space: nowrap; }
.toolbar input[type="text"] { flex: 1; min-width: 120px; padding: 5px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px; }
.toolbar input[type="file"] { display: none; }
.btn { padding: 6px 14px; border: 1px solid #bbb; border-radius: 3px; cursor: pointer; font-size: 13px; background: #e8e8e8; color: #222; white-space: nowrap; }
.btn:hover { background: #ddd; }
.btn-primary { background: #4a90d9; color: #fff; border-color: #357abd; font-weight: 600; }
.btn-primary:hover { background: #357abd; }

/* Stats bar */
.stats { padding: 4px 8px; background: #e8e8e8; font-size: 12px; color: #555; min-height: 22px; }

/* Diff panels */
.diff-container { display: flex; flex: 1; gap: 6px; overflow: hidden; min-height: 0; }
.diff-panel { flex: 1; display: flex; flex-direction: column; overflow: hidden; border: 1px solid #ccc; border-radius: 3px; background: #fff; }
.diff-panel-header { background: #d0d0d0; padding: 4px 8px; font-weight: 600; font-size: 13px; }
.diff-panel-body { flex: 1; overflow-y: auto; overflow-x: hidden; padding: 6px 4px; font-family: var(--diff-font, 'Consolas', 'Courier New', 'Microsoft YaHei', 'SimSun', monospace); font-size: var(--diff-font-size, 13px); line-height: 1.6; white-space: pre-wrap; word-break: break-all; }

/* Manual input */
.manual-inputs { display: flex; gap: 6px; padding: 6px; background: #f0f0f0; flex: 0 0 auto; }
.manual-inputs textarea { flex: 1; padding: 6px; border: 1px solid #ccc; border-radius: 3px; font-family: 'Consolas', 'Courier New', 'Microsoft YaHei', monospace; font-size: 13px; resize: vertical; min-height: 90px; max-height: 200px; }

/* Color settings */
.color-grid { display: grid; grid-template-columns: 220px 70px auto auto; gap: 4px; align-items: center; padding: 8px; background: #f0f0f0; }
.color-grid-header { font-weight: 600; font-size: 12px; color: #555; padding: 2px 4px; background: #d0d0d0; }
.color-preview { width: 64px; height: 22px; border: 1px solid #aaa; border-radius: 2px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; }
.color-picker-btn { padding: 3px 8px; border: 1px solid #aaa; border-radius: 2px; cursor: pointer; font-size: 12px; background: #e8e8e8; text-align: center; }
.color-actions { padding: 8px; display: flex; gap: 6px; align-items: center; }

/* Diff tags */
.tag-equal {}
.tag-added { background: #c8f7c5; color: #1b5e20; }
.tag-removed { background: #f8d7da; color: #7f1d1d; }
.tag-changed { background: #fff3cd; color: #664d03; }
.tag-inline-add { background: #81c784; color: #1b5e20; }
.tag-inline-remove { background: #e57373; color: #7f1d1d; }
.line-num { color: #999; font-size: 11px; user-select: none; }
.diff-line { padding: 0 4px; border-radius: 1px; }
.placeholder-line { visibility: hidden; white-space: pre-wrap; word-break: break-all; }
.diff-row { white-space: pre-wrap; word-break: break-all; }

/* Scroll sync */
.diff-panel-body::-webkit-scrollbar { width: 8px; }
.diff-panel-body::-webkit-scrollbar-track { background: #f1f1f1; }
.diff-panel-body::-webkit-scrollbar-thumb { background: #bbb; border-radius: 4px; }

input[type="color"] { width: 60px; height: 28px; padding: 0; border: 1px solid #aaa; border-radius: 2px; cursor: pointer; background: none; }

/* Status messages */
.status-ok { color: #1b5e20; padding: 20px; text-align: center; }
.status-warn { color: #856404; padding: 4px 8px; }
</style>
</head>
<body>

<div class="tab-bar">
  <button class="tab-btn active" onclick="switchTab('file')">File Compare</button>
  <button class="tab-btn" onclick="switchTab('manual')">Manual Compare</button>
  <button class="tab-btn" onclick="switchTab('color')">Color Settings</button>
  <button class="tab-btn" onclick="switchTab('font')">Font Settings</button>
  <button class="tab-btn" onclick="shutdownServer()" style="margin-left:auto;background:#c0392b;color:#fff;border:none">✕ Exit</button>
</div>

<!-- ─── File Compare Tab ─── -->
<div id="tab-file" class="tab-content active">
  <div class="toolbar">
    <label>File 1:</label>
    <input type="text" id="fc-file1" placeholder="Click Browse..." readonly>
    <input type="file" id="fc-file1-input" accept=".docx" onchange="loadFile(1, this)">
    <button class="btn" onclick="document.getElementById('fc-file1-input').click()">Browse...</button>
    <label style="margin-left:8px">File 2:</label>
    <input type="text" id="fc-file2" placeholder="Click Browse..." readonly>
    <input type="file" id="fc-file2-input" accept=".docx" onchange="loadFile(2, this)">
    <button class="btn" onclick="document.getElementById('fc-file2-input').click()">Browse...</button>
    <button class="btn btn-primary" onclick="doFileCompare()">Compare</button>
    <button class="btn" onclick="exportResult()">Export</button>
  </div>
  <div class="stats" id="fc-stats">Select two files and click Compare</div>
  <div class="diff-container" id="fc-diff-container">
    <div class="diff-panel">
      <div class="diff-panel-header">File 1 (Original)</div>
      <div class="diff-panel-body" id="fc-left"></div>
    </div>
    <div class="diff-panel">
      <div class="diff-panel-header">File 2 (Modified)</div>
      <div class="diff-panel-body" id="fc-right"></div>
    </div>
  </div>
</div>

<!-- ─── Manual Compare Tab ─── -->
<div id="tab-manual" class="tab-content">
  <div class="manual-inputs">
    <textarea id="mc-left" placeholder="Original text (paste here)..."></textarea>
    <textarea id="mc-right" placeholder="Modified text (paste here)..."></textarea>
  </div>
  <div class="toolbar">
    <button class="btn btn-primary" onclick="doManualCompare()">Compare Text</button>
    <button class="btn" onclick="clearManual()">Clear</button>
    <span class="stats" id="mc-stats" style="margin-left:8px;display:inline-block;background:transparent"></span>
  </div>
  <div class="diff-container" style="flex:1">
    <div class="diff-panel">
      <div class="diff-panel-header">Original</div>
      <div class="diff-panel-body" id="mc-left-panel"></div>
    </div>
    <div class="diff-panel">
      <div class="diff-panel-header">Modified</div>
      <div class="diff-panel-body" id="mc-right-panel"></div>
    </div>
  </div>
</div>

<!-- ─── Color Settings Tab ─── -->
<div id="tab-color" class="tab-content" style="overflow-y:auto">
  <div style="padding:12px"><h3 style="margin-bottom:8px">Diff Highlight Colors</h3></div>
  <div class="color-grid" id="color-grid">
    <div class="color-grid-header">Diff Element</div>
    <div class="color-grid-header">Preview</div>
    <div class="color-grid-header">Foreground</div>
    <div class="color-grid-header">Background</div>
  </div>
  <div class="color-actions">
    <button class="btn" onclick="resetColors()">Reset Defaults</button>
    <button class="btn" onclick="importColors()">Import Scheme...</button>
    <button class="btn" onclick="exportColors()">Export Scheme...</button>
    <button class="btn btn-primary" onclick="applyColors()">Apply</button>
    <input type="file" id="color-import-input" accept=".json" style="display:none" onchange="doImportColors(this)">
  </div>
  <div id="color-status" class="stats" style="margin-top:4px"></div>
</div>

<!-- ─── Font Settings Tab ─── -->
<div id="tab-font" class="tab-content" style="overflow-y:auto">
  <div style="padding:12px">
    <h3 style="margin-bottom:8px">Font Settings</h3>
    <p style="font-size:12px;color:#666;margin-bottom:12px">Press <kbd style="border:1px solid #aaa;border-radius:2px;padding:1px 5px;background:#eee">F5</kbd> after changing to re-render.</p>
    <div style="display:flex;flex-direction:column;gap:12px;max-width:480px">
      <div>
        <label style="font-weight:500">Font Family:</label>
        <select id="font-family" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:3px;margin-top:4px;font-size:13px">
          <option value="'Consolas','Courier New','Microsoft YaHei','SimSun',monospace">Consolas / Microsoft YaHei (Default)</option>
          <option value="'Courier New','SimSun',monospace">Courier New / SimSun</option>
          <option value="'Microsoft YaHei','SimHei',sans-serif">Microsoft YaHei</option>
          <option value="'SimSun','SimHei',serif">SimSun (宋体)</option>
          <option value="'SimHei',sans-serif">SimHei (黑体)</option>
          <option value="'Fira Code','Courier New',monospace">Fira Code</option>
          <option value="'JetBrains Mono','Courier New',monospace">JetBrains Mono</option>
          <option value="'Noto Sans SC',sans-serif">Noto Sans SC</option>
          <option value="sans-serif">System Sans</option>
          <option value="monospace">System Monospace</option>
        </select>
      </div>
      <div>
        <label style="font-weight:500">Font Size:</label>
        <div style="display:flex;gap:8px;align-items:center;margin-top:4px">
          <input type="range" id="font-size-slider" min="10" max="24" value="13" step="1"
                 style="flex:1" oninput="document.getElementById('font-size-label').textContent=this.value+'px'">
          <span id="font-size-label" style="font-size:13px;min-width:36px">13px</span>
          <input type="number" id="font-size-input" value="13" min="10" max="24" step="1"
                 style="width:60px;padding:4px;border:1px solid #ccc;border-radius:3px;text-align:center;font-size:13px"
                 onchange="document.getElementById('font-size-slider').value=this.value;document.getElementById('font-size-label').textContent=this.value+'px'">
        </div>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn btn-primary" onclick="applyFontSettings()">Apply</button>
        <button class="btn" onclick="resetFontSettings()">Reset</button>
      </div>
      <div id="font-status" class="stats" style="font-size:12px">Current: 13px / Default monospace</div>
    </div>
  </div>
</div>

<script>
// ────────────────────── State ──────────────────────
const COLORS = {
    equal:      { fg: "#000000", bg: "#ffffff" },
    added:      { fg: "#1b5e20", bg: "#c8f7c5" },
    removed:    { fg: "#7f1d1d", bg: "#f8d7da" },
    changed:    { fg: "#664d03", bg: "#fff3cd" },
    inline_add: { fg: "#1b5e20", bg: "#81c784" },
    inline_rm:  { fg: "#7f1d1d", bg: "#e57373" },
};

const TAG_LABELS = {
    equal: "Equal (same text)",
    added: "Added (new text)",
    removed: "Removed (deleted text)",
    changed: "Changed (modified text)",
    inline_add: "Inline added (char-level)",
    inline_rm: "Inline removed (char-level)",
};

let fcLeftParas = [];
let fcRightParas = [];
let fcDiffResult = [];
let mcDiffResult = [];
let syncScrolling = false;
const SAVE_NAME = "docx_compare_colors";
const FONT_SAVE_NAME = "docx_compare_font";

// ────────────────────── Font Settings ──────────────────────
function initFontSettings() {
    const saved = localStorage.getItem(FONT_SAVE_NAME);
    if (saved) {
        try {
            const data = JSON.parse(saved);
            document.getElementById('font-family').value = data.family;
            document.getElementById('font-size-slider').value = data.size;
            document.getElementById('font-size-input').value = data.size;
            document.getElementById('font-size-label').textContent = data.size + 'px';
            applyFontVars(data.family, data.size);
        } catch(e) {}
    }
}

function applyFontSettings() {
    const family = document.getElementById('font-family').value;
    const size = document.getElementById('font-size-input').value;
    applyFontVars(family, size);
    localStorage.setItem(FONT_SAVE_NAME, JSON.stringify({family, size}));
    document.getElementById('font-status').textContent =
        `Applied: ${size}px / ${family.slice(0, 40)}…`;
    // Re-render if results exist
    if (fcDiffResult.length) renderDiff('fc', fcDiffResult);
    if (mcDiffResult.length) renderDiff('mc', mcDiffResult);
}

function applyFontVars(family, size) {
    document.documentElement.style.setProperty('--diff-font', family);
    document.documentElement.style.setProperty('--diff-font-size', size + 'px');
}

function resetFontSettings() {
    const defFamily = "'Consolas','Courier New','Microsoft YaHei','SimSun',monospace";
    const defSize = 13;
    document.getElementById('font-family').value = defFamily;
    document.getElementById('font-size-slider').value = defSize;
    document.getElementById('font-size-input').value = defSize;
    document.getElementById('font-size-label').textContent = defSize + 'px';
    applyFontVars(defFamily, defSize);
    localStorage.removeItem(FONT_SAVE_NAME);
    document.getElementById('font-status').textContent = 'Reset to defaults.';
    if (fcDiffResult.length) renderDiff('fc', fcDiffResult);
    if (mcDiffResult.length) renderDiff('mc', mcDiffResult);
}

function shutdownServer() {
    fetch('/api/shutdown', {method: 'POST'}).then(() => {
        document.body.innerHTML = '<div style="padding:40px;text-align:center"><h2>Server shutdown.</h2><p>You may close this tab.</p></div>';
    }).catch(() => {
        document.body.innerHTML = '<div style="padding:40px;text-align:center"><h2>Server stopped.</h2><p>You may close this tab.</p></div>';
    });
}

// ────────────────────── Tab Switch ──────────────────────
function switchTab(name) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    document.querySelector(`.tab-btn[onclick*="'${name}'"]`).classList.add('active');
    document.getElementById(`tab-${name}`).classList.add('active');
}

// ────────────────────── Color Management ──────────────────────
function initColorGrid() {
    const grid = document.getElementById('color-grid');
    const order = ['equal','added','removed','changed','inline_add','inline_rm'];
    order.forEach(tag => {
        const c = COLORS[tag];
        grid.innerHTML +=
            `<div style="padding:2px 4px;font-size:13px">${TAG_LABELS[tag]}</div>` +
            `<div class="color-preview" id="preview-${tag}" style="background:${c.bg};color:${c.fg}">Aa</div>` +
            `<input type="color" id="fg-${tag}" value="${c.fg}" onchange="onColorChange('${tag}')">` +
            `<input type="color" id="bg-${tag}" value="${c.bg}" onchange="onColorChange('${tag}')">`;
    });
}

function onColorChange(tag) {
    const fg = document.getElementById(`fg-${tag}`).value;
    const bg = document.getElementById(`bg-${tag}`).value;
    const prev = document.getElementById(`preview-${tag}`);
    prev.style.color = fg;
    prev.style.background = bg;
}

function resetColors() {
    const def = {equal:{fg:"#000000",bg:"#ffffff"},added:{fg:"#1b5e20",bg:"#c8f7c5"},removed:{fg:"#7f1d1d",bg:"#f8d7da"},changed:{fg:"#664d03",bg:"#fff3cd"},inline_add:{fg:"#1b5e20",bg:"#81c784"},inline_rm:{fg:"#7f1d1d",bg:"#e57373"}};
    Object.keys(def).forEach(tag => {
        document.getElementById(`fg-${tag}`).value = def[tag].fg;
        document.getElementById(`bg-${tag}`).value = def[tag].bg;
        document.getElementById(`preview-${tag}`).style.color = def[tag].fg;
        document.getElementById(`preview-${tag}`).style.background = def[tag].bg;
        COLORS[tag] = {fg: def[tag].fg, bg: def[tag].bg};
    });
    document.getElementById('color-status').textContent = 'Reset to defaults. Click Apply to use.';
}

function applyColors() {
    Object.keys(COLORS).forEach(tag => {
        COLORS[tag] = {
            fg: document.getElementById(`fg-${tag}`).value,
            bg: document.getElementById(`bg-${tag}`).value,
        };
    });
    localStorage.setItem(SAVE_NAME, JSON.stringify(COLORS));
    // Re-render both panels
    if (fcDiffResult.length) renderDiff('fc', fcDiffResult);
    if (mcDiffResult.length) renderDiff('mc', mcDiffResult);
    document.getElementById('color-status').textContent = 'Colors applied!';
}

function importColors() {
    document.getElementById('color-import-input').click();
}

function doImportColors(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const data = JSON.parse(e.target.result);
            Object.keys(data).forEach(tag => {
                if (COLORS[tag]) {
                    const fg = data[tag].fg || COLORS[tag].fg;
                    const bg = data[tag].bg || COLORS[tag].bg;
                    document.getElementById(`fg-${tag}`).value = fg;
                    document.getElementById(`bg-${tag}`).value = bg;
                    document.getElementById(`preview-${tag}`).style.color = fg;
                    document.getElementById(`preview-${tag}`).style.background = bg;
                    COLORS[tag] = {fg, bg};
                }
            });
            document.getElementById('color-status').textContent = `Imported from ${file.name}`;
        } catch(err) {
            alert('Invalid color scheme file: ' + err.message);
        }
    };
    reader.readAsText(file);
    input.value = '';
}

function exportColors() {
    const data = {};
    Object.keys(COLORS).forEach(tag => {
        data[tag] = {
            fg: document.getElementById(`fg-${tag}`).value,
            bg: document.getElementById(`bg-${tag}`).value,
        };
    });
    const blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'color_scheme.json';
    a.click();
    URL.revokeObjectURL(a.href);
}

function loadSavedColors() {
    try {
        const saved = localStorage.getItem(SAVE_NAME);
        if (saved) {
            const data = JSON.parse(saved);
            Object.keys(data).forEach(tag => {
                if (COLORS[tag]) {
                    COLORS[tag] = data[tag];
                    const el_fg = document.getElementById(`fg-${tag}`);
                    const el_bg = document.getElementById(`bg-${tag}`);
                    const el_pr = document.getElementById(`preview-${tag}`);
                    if (el_fg) el_fg.value = data[tag].fg;
                    if (el_bg) el_bg.value = data[tag].bg;
                    if (el_pr) { el_pr.style.color = data[tag].fg; el_pr.style.background = data[tag].bg; }
                }
            });
        }
    } catch(e) {}
}

// ────────────────────── File Compare ──────────────────────
async function loadFile(side, input) {
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = async function(e) {
        const content = e.target.result.split(',')[1]; // data:...;base64,XXXX -> XXXX
        const resp = await fetch('/api/parse_docx', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({filename: file.name, content: content })
        });
        const result = await resp.json();
        const paras = result.paras;
        document.getElementById(`fc-file${side}`).value = file.name;
        if (side === 1) fcLeftParas = paras; else fcRightParas = paras;
    };
    reader.readAsDataURL(file);
}

async function doFileCompare() {
    const f1 = document.getElementById('fc-file1').value;
    const f2 = document.getElementById('fc-file2').value;
    if (!f1 || !f2) { alert('Please select both files.'); return; }

    // Filter empty paragraphs client-side for stats accuracy
    const leftFiltered = fcLeftParas.filter(p => p.trim());
    const rightFiltered = fcRightParas.filter(p => p.trim());

    const resp = await fetch('/api/compare', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ left: fcLeftParas, right: fcRightParas })
    });
    const result = await resp.json();
    fcDiffResult = result.diff;

    let eq=0, ad=0, rm=0, rp=0;
    result.diff.forEach(d => { if(d[0]==='equal') eq++; else if(d[0]==='add') ad++; else if(d[0]==='remove') rm++; else if(d[0]==='replace') rp++; });
    document.getElementById('fc-stats').textContent =
        `File1: ${leftFiltered.length} paras | File2: ${rightFiltered.length} paras | Equal: ${eq}  Added: ${ad}  Removed: ${rm}  Modified: ${rp}`;

    renderDiff('fc', result.diff);
}

// ────────────────────── Manual Compare ──────────────────────
async function doManualCompare() {
    const left = document.getElementById('mc-left').value.trim();
    const right = document.getElementById('mc-right').value.trim();
    if (!left && !right) { alert('Please enter text in at least one field.'); return; }

    const leftLines = left ? left.split('\n').filter(l => l.trim()).map(l => l.trim()) : [];
    const rightLines = right ? right.split('\n').filter(l => l.trim()).map(l => l.trim()) : [];

    const resp = await fetch('/api/compare', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ left: leftLines, right: rightLines })
    });
    const result = await resp.json();
    mcDiffResult = result.diff;

    let eq=0, ad=0, rm=0, rp=0;
    result.diff.forEach(d => { if(d[0]==='equal') eq++; else if(d[0]==='add') ad++; else if(d[0]==='remove') rm++; else if(d[0]==='replace') rp++; });
    document.getElementById('mc-stats').textContent =
        `Original: ${leftLines.length} lines | Modified: ${rightLines.length} lines | Equal: ${eq}  Added: ${ad}  Removed: ${rm}  Modified: ${rp}`;

    renderDiff('mc', result.diff);
}

function clearManual() {
    document.getElementById('mc-left').value = '';
    document.getElementById('mc-right').value = '';
    document.getElementById('mc-left-panel').innerHTML = '';
    document.getElementById('mc-right-panel').innerHTML = '';
    document.getElementById('mc-stats').textContent = '';
    mcDiffResult = [];
}

// ────────────────────── Render Diff ──────────────────────
function renderDiff(prefix, diff) {
    const leftPanel = document.getElementById(prefix + '-left-panel') || document.getElementById(prefix + '-left');
    const rightPanel = document.getElementById(prefix + '-right-panel') || document.getElementById(prefix + '-right');
    leftPanel.innerHTML = '';
    rightPanel.innerHTML = '';

    let n = 0;
    diff.forEach(([tag, lt, rt]) => {
        n++;
        const lineNum = `<span class="line-num">${String(n).padStart(3)} | </span>`;
        const rowId = `r${prefix}-${n}`;

        // Phase 1: build row HTML
        let leftHtml = '', rightHtml = '';
        if (tag === 'equal') {
            leftHtml = lineNum + escHtml(lt);
            rightHtml = lineNum + escHtml(rt);
        } else if (tag === 'add') {
            leftHtml = lineNum + `<span class="placeholder-line">${escHtml(rt)}</span>`;
            rightHtml = lineNum + `<span class="tag-added">+ </span><span class="tag-added">${escHtml(rt)}</span>`;
        } else if (tag === 'remove') {
            leftHtml = lineNum + `<span class="tag-removed">- </span><span class="tag-removed">${escHtml(lt)}</span>`;
            rightHtml = lineNum + `<span class="placeholder-line">${escHtml(lt)}</span>`;
        } else if (tag === 'replace') {
            const sm = inlineDiffHtml(lt || '', rt || '');
            leftHtml = lineNum + `<span class="tag-changed">~ </span>` + sm.left;
            rightHtml = lineNum + `<span class="tag-changed">~ </span>` + sm.right;
        }
        // Wrap in a row div for height measurement
        leftPanel.innerHTML += `<div class="diff-row" data-row="${n}">${leftHtml}</div>`;
        rightPanel.innerHTML += `<div class="diff-row" data-row="${n}">${rightHtml}</div>`;
    });

    // Phase 2: balance heights — set min-height on the shorter row to match the taller one
    // Clear any previous min-heights first, wait for layout to settle, then measure & set
    const allRows = leftPanel.querySelectorAll('.diff-row');
    allRows.forEach(r => r.style.minHeight = '');
    const allRowsR = rightPanel.querySelectorAll('.diff-row');
    allRowsR.forEach(r => r.style.minHeight = '');

    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            for (let i = 1; i <= diff.length; i++) {
                const lRow = leftPanel.querySelector(`.diff-row[data-row="${i}"]`);
                const rRow = rightPanel.querySelector(`.diff-row[data-row="${i}"]`);
                if (!lRow || !rRow) continue;
                // Force layout reflow
                void lRow.offsetHeight;
                void rRow.offsetHeight;
                const lh = lRow.getBoundingClientRect().height;
                const rh = rRow.getBoundingClientRect().height;
                if (Math.abs(lh - rh) < 0.5) continue;  // tolerance for sub-pixel
                if (lh < rh) {
                    lRow.style.minHeight = rh + 'px';
                } else {
                    rRow.style.minHeight = lh + 'px';
                }
            }
        });
    });
}

function inlineDiffHtml(lt, rt) {
    // Simple char-level diff rendered in HTML
    let leftHtml = '', rightHtml = '';
    let i = 0, j = 0;
    // Use LCS-like approach for display
    const lcs = longestCommonSubsequence(lt, rt);
    let li = 0, ri = 0;
    for (const [lc, rc] of lcs) {
        // Removed chars in left before match
        if (li < lc) {
            leftHtml += `<span class="tag-inline-remove">${escHtml(lt.slice(li, lc))}</span>`;
            li = lc;
        }
        // Added chars in right before match
        if (ri < rc) {
            rightHtml += `<span class="tag-inline-add">${escHtml(rt.slice(ri, rc))}</span>`;
            ri = rc;
        }
        // Equal part
        let eqLen = 1;
        while (li + eqLen <= lt.length && ri + eqLen <= rt.length &&
               lt[li + eqLen - 1] === rt[ri + eqLen - 1]) eqLen++;
        eqLen--;
        if (eqLen > 0) {
            leftHtml += escHtml(lt.slice(li, li + eqLen));
            rightHtml += escHtml(rt.slice(ri, ri + eqLen));
            li += eqLen;
            ri += eqLen;
        }
    }
    // Remaining
    if (li < lt.length) leftHtml += `<span class="tag-inline-remove">${escHtml(lt.slice(li))}</span>`;
    if (ri < rt.length) rightHtml += `<span class="tag-inline-add">${escHtml(rt.slice(ri))}</span>`;
    return {left: leftHtml, right: rightHtml};
}

function longestCommonSubsequence(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({length: m+1}, () => new Int32Array(n+1));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] + 1 : Math.max(dp[i-1][j], dp[i][j-1]);
        }
    }
    // Backtrack
    const result = [];
    let i = m, j = n;
    while (i > 0 && j > 0) {
        if (a[i-1] === b[j-1]) {
            result.unshift([i-1, j-1]);
            i--; j--;
        } else if (dp[i-1][j] > dp[i][j-1]) {
            i--;
        } else {
            j--;
        }
    }
    return result;
}

// ────────────────────── Export ──────────────────────
async function exportResult() {
    if (!fcDiffResult.length) { alert('Please run comparison first.'); return; }

    const resp = await fetch('/api/export', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            diff: fcDiffResult,
            stats: {
                left_paras: fcLeftParas.length,
                right_paras: fcRightParas.length,
            }
        })
    });
    const blob = await resp.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const now = new Date();
    a.download = `docx_diff_${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}${String(now.getSeconds()).padStart(2,'0')}.txt`;
    a.click();
    URL.revokeObjectURL(a.href);
}

// ────────────────────── Sync Scroll ──────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initColorGrid();
    loadSavedColors();
    initFontSettings();

    // F5 — refresh + re-run comparison
    document.addEventListener('keydown', e => {
        if (e.key === 'F5') {
            e.preventDefault();
            // Re-render both compare panels
            if (fcDiffResult.length) renderDiff('fc', fcDiffResult);
            if (mcDiffResult.length) renderDiff('mc', mcDiffResult);
        }
    });

    ['fc', 'mc'].forEach(prefix => {
        const left = document.getElementById(prefix + '-left-panel') || document.getElementById(prefix + '-left');
        const right = document.getElementById(prefix + '-right-panel') || document.getElementById(prefix + '-right');
        if (!left || !right) return;
        left.addEventListener('scroll', () => {
            if (!syncScrolling) { syncScrolling = true; right.scrollTop = left.scrollTop; right.scrollLeft = left.scrollLeft; setTimeout(() => syncScrolling = false, 50); }
        });
        right.addEventListener('scroll', () => {
            if (!syncScrolling) { syncScrolling = true; left.scrollTop = right.scrollTop; left.scrollLeft = right.scrollLeft; setTimeout(() => syncScrolling = false, 50); }
        });
    });
});

// ────────────────────── Utils ──────────────────────
function escHtml(s) {
    if (!s) return '';
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>"""


# ──────────────────────── HTTP Server ────────────────────────
class DocxCompareHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers['Content-Length'])
        body = json.loads(self.rfile.read(length))

        if self.path == '/api/parse_docx':
            self._handle_parse(body)
        elif self.path == '/api/compare':
            self._handle_compare(body)
        elif self.path == '/api/export':
            self._handle_export(body)
        elif self.path == '/api/shutdown':
            self._handle_shutdown()
        else:
            self.send_response(404)
            self.end_headers()

    def _send_json(self, data):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))

    def _handle_parse(self, body):
        filename = body.get('filename', '')
        content_b64 = body.get('content', '')
        import base64
        content = base64.b64decode(content_b64)
        # Save to temp file
        tmp_path = os.path.join(temp_dir, filename)
        with open(tmp_path, 'wb') as f:
            f.write(content)
        try:
            paras = DocxExtractor.extract_text(tmp_path)
            self._send_json({"paras": paras})
        except Exception as e:
            self._send_json({"paras": [], "error": str(e)})

    def _handle_compare(self, body):
        left = body.get('left', [])
        right = body.get('right', [])
        diff = DiffEngine.compare_paragraphs(left, right)
        self._send_json({"diff": diff})

    def _handle_export(self, body):
        diff = body.get('diff', [])
        stats = body.get('stats', {})
        lines = [
            "DOCX Comparison Result",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 60,
            f"File1: {stats.get('left_paras', 0)} paras | File2: {stats.get('right_paras', 0)} paras",
            "=" * 60, ""
        ]
        n = 0
        for tag, lt, rt in diff:
            n += 1
            if tag == "equal":
                lines.append(f"[{n}] (equal) {lt}")
            elif tag == "add":
                lines.append(f"[{n}] + added: {rt}")
            elif tag == "remove":
                lines.append(f"[{n}] - removed: {lt}")
            elif tag == "replace":
                lines.append(f"[{n}] ~ modified:")
                lines.append(f"    original: {lt}")
                lines.append(f"    changed to: {rt}")
            lines.append("")
        text = "\n".join(lines)
        self.send_response(200)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.send_header('Content-Disposition', 'attachment')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(text.encode('utf-8'))

    def _handle_shutdown(self):
        self.send_response(200)
        self.send_header('Content-Type', 'text/plain')
        self.end_headers()
        self.wfile.write(b'Shutting down...')
        # Schedule shutdown in a thread to avoid deadlock
        import threading
        threading.Thread(target=self.server.shutdown, daemon=True).start()

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            with open(HTML_FILE, 'rb') as f:
                self.wfile.write(f.read())
        else:
            super().do_GET()

    def log_message(self, format, *args):
        pass  # suppress logs


def main():
    global HTML_FILE, temp_dir
    import tempfile
    temp_dir = tempfile.mkdtemp(prefix="docx_compare_")

    # Write HTML file
    html_content = generate_html()
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html_content)

    # Start server
    server = socketserver.TCPServer(("", PORT), DocxCompareHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    url = f"http://localhost:{PORT}"
    bar = "=" * 45
    print(f"\n  📄 DOCX Compare Tool - Web Edition")
    print(f"  {bar}")
    print(f"  Opening browser at: {url}")
    print(f"  Press Ctrl+C to stop the server\n")
    webbrowser.open(url)

    try:
        while True:
            import time
            time.sleep(3600)
    except KeyboardInterrupt:
        print("\nShutting down...")
        server.shutdown()
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        if os.path.exists(HTML_FILE):
            os.remove(HTML_FILE)


if __name__ == "__main__":
    main()
