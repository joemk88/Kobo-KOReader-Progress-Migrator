import os
import re
import shutil
import sqlite3
import datetime
import threading
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

BG       = "#1c1c1e"
PANEL    = "#2c2c2e"
PANEL2   = "#3a3a3c"
ACCENT   = "#0a84ff"
GREEN    = "#30d158"
RED      = "#ff453a"
AMBER    = "#ffd60a"
TEXT     = "#f2f2f7"
MUTED    = "#8e8e93"
ENTRY_BG = "#1c1c1e"
BORDER   = "#48484a"
FONT_MONO  = ("Consolas", 10)
FONT_UI    = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_TITLE = ("Segoe UI Semibold", 11)
FONT_HEAD  = ("Segoe UI Semibold", 14)
APPIMAGE_NAME = "koreader-v2026.03-x86_64.AppImage"

# ---------------- path utils ----------------

def win_to_wsl(path: str) -> str:
    path = os.path.abspath(path)
    if len(path) >= 2 and path[1] == ":":
        drive = path[0].lower()
        rest = path[2:].replace("\\", "/")
        if rest.startswith("/"):
            return f"/mnt/{drive}{rest}"
        return f"/mnt/{drive}/{rest}"
    return path.replace("\\", "/")


def parse_settings_lua(settings_path):
    result = {}
    try:
        txt = open(settings_path, encoding="utf-8", errors="replace").read()
        for key in ("home_dir", "lastdir"):
            m = re.search(rf'\["{key}"\]\s*=\s*"([^"]+)"', txt)
            if m:
                result[key] = m.group(1)
    except Exception:
        pass
    return result


def find_settings_lua(koreader_folder):
    p = os.path.join(koreader_folder, "settings.reader.lua")
    return p if os.path.isfile(p) else None


def derive_device_books_root(home_dir, local_books_folder):
    home_dir = (home_dir or "").rstrip("/")
    if not home_dir:
        return ""
    base = os.path.basename(os.path.normpath(local_books_folder or ""))
    if base:
        return f"{home_dir}/{base}"
    return home_dir



def get_launcher_dir():
    return os.path.dirname(os.path.abspath(__file__))


def get_manual_output_root():
    return os.path.join(get_launcher_dir(), "OUTPUT")


def get_manual_sidecar_path(book_path, books_folder):
    rel = os.path.relpath(book_path, books_folder)
    rel_no_ext, _ = os.path.splitext(rel)
    return os.path.join(get_manual_output_root(), "BOOKS", rel_no_ext + ".sdr", metadata_lua_name(book_path))


def write_manual_instructions(dev_books, manual_output_root, log_fn=None):
    books_out = os.path.join(manual_output_root, "BOOKS")
    koreader_out = os.path.join(manual_output_root, "KOREADER")
    instructions = f"""KOReader Manual Transfer Instructions

This run was completed in Manual Transfer Mode.
No files were written directly to your target device.

1. Copy everything INSIDE this folder:
   {books_out}

   into your target device's Books root:
   {dev_books}

   Important: preserve the folder structure exactly.
   The .sdr folders have been created in the same relative paths as the matching book files.

2. Copy this file:
   {os.path.join(koreader_out, 'history.lua')}

   into your target KOReader folder as:
   history.lua

   Overwrite the existing file if prompted.

3. Restart KOReader on the target device.

Tip:
- If your books are organised by author / series / subfolders, keep that structure exactly when pasting.
- If your books are all loose in one folder, paste the .sdr folders alongside those books in that same folder.
"""
    readme = """Manual Transfer Mode

The OUTPUT folder contains files ready to copy to a target device when the target storage is unreliable (for example Android Shared Internal Storage, CrossDevice, or unstable MTP).

- OUTPUT/BOOKS mirrors the relative structure of the selected Books folder.
- Each generated .sdr folder belongs next to its matching book file.
- OUTPUT/KOREADER/history.lua should be copied into the target KOReader folder.

You can usually complete the transfer by opening OUTPUT/BOOKS, copying all contents, and pasting them into the target Books root. Then copy OUTPUT/KOREADER/history.lua into the target KOReader folder.
"""
    os.makedirs(manual_output_root, exist_ok=True)
    with open(os.path.join(manual_output_root, 'INSTRUCTIONS.txt'), 'w', encoding='utf-8') as f:
        f.write(instructions)
    with open(os.path.join(manual_output_root, 'README_COPY.txt'), 'w', encoding='utf-8') as f:
        f.write(readme)
    if log_fn:
        log_fn('INFO', f'Manual transfer output written to: {manual_output_root}')
        log_fn('INFO', 'Read INSTRUCTIONS.txt for copy/paste steps.')

# ---------------- kobo db ----------------

def read_kobo_books(sqlite_path):
    conn = sqlite3.connect(f"file:{sqlite_path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT Title, Attribution, ___PercentRead, ReadStatus,
               DateLastRead, ContentID, ChapterIDBookmarked,
               ParagraphBookmarked, BookmarkWordOffset, ___NumPages
        FROM content
        WHERE ContentType = 6
          AND ReadStatus IN (1, 2)
          AND Title IS NOT NULL AND Title != ''
        ORDER BY DateLastRead DESC
    """)
    rows = cur.fetchall()
    conn.close()
    books = []
    for row in rows:
        pct = float(row["___PercentRead"] or 0)
        status_id = row["ReadStatus"] or 0
        status = "complete" if status_id == 2 else "reading"
        percent = 1.0 if status_id == 2 else pct / 100.0
        cid = row["ContentID"] or ""
        hint = cid.split("://", 1)[-1] if "://" in cid else cid
        chapter = row["ChapterIDBookmarked"] or ""
        chapter_file = chapter.split("#", 1)[0].lstrip("/") if chapter else ""
        books.append({
            "title": row["Title"] or "",
            "author": row["Attribution"] or "",
            "percent": percent,
            "status": status,
            "last_read": row["DateLastRead"] or "",
            "content_id": cid,
            "chapter_id": chapter,
            "chapter_file": chapter_file,
            "paragraph_bookmarked": row["ParagraphBookmarked"] or 0,
            "bookmark_word_offset": row["BookmarkWordOffset"] or 0,
            "file_basename": os.path.basename(hint),
            "num_pages": row["___NumPages"] or 0,
        })
    return books

# ---------------- matching ----------------

def norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def scan_books(folder):
    exts = {".epub", ".pdf", ".mobi", ".cbz", ".cbr"}
    hits = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.endswith(".sdr")]
        for f in files:
            lower = f.lower()
            if lower.endswith(".kepub.epub"):
                hits.append((f, os.path.join(root, f)))
                continue
            if os.path.splitext(lower)[1] in exts:
                hits.append((f, os.path.join(root, f)))
    return hits


def match_books(kobo_books, files):
    by_name = {norm(f): p for f, p in files}
    results = []
    for book in kobo_books:
        path = None
        reason = "no match"
        nb = norm(book["file_basename"])
        if nb and nb in by_name:
            path = by_name[nb]
            reason = "filename"
        if not path and book["title"]:
            nt = norm(book["title"])
            for fname, fpath in files:
                if nt and nt in norm(fname):
                    path = fpath
                    reason = "title"
                    break
        results.append((book, path, reason))
    return results

# ---------------- sidecar utils ----------------

def metadata_lua_name(book_path):
    ext = os.path.splitext(book_path)[1].lower().lstrip(".") or "epub"
    return f"metadata.{ext}.lua"


def get_sidecar_path(book_path):
    root, _ = os.path.splitext(book_path)
    return os.path.join(root + ".sdr", metadata_lua_name(book_path))


def remove_old_metadata_backups(lua_path):
    sdr = os.path.dirname(lua_path)
    removed = []
    if os.path.isdir(sdr):
        for name in os.listdir(sdr):
            if name.startswith("metadata.") and name.endswith(".lua.old"):
                full = os.path.join(sdr, name)
                try:
                    os.remove(full)
                    removed.append(full)
                except Exception:
                    pass
    return removed


def patch_or_insert(txt, key, replacement_line):
    pattern = rf'\["{re.escape(key)}"\]\s*=\s*.*?,(\s*)'
    if re.search(pattern, txt, flags=re.DOTALL):
        return re.sub(pattern, replacement_line + r',\1', txt, count=1, flags=re.DOTALL), True
    txt = re.sub(r'(\}\s*)$', f'    {replacement_line},\n\\1', txt, count=1, flags=re.DOTALL)
    return txt, False


def patch_last_page(txt, percent, fallback_pages=0):
    doc_pages = None
    m = re.search(r'\["doc_pages"\]\s*=\s*(\d+)', txt)
    if m:
        doc_pages = int(m.group(1))
    elif fallback_pages:
        doc_pages = int(fallback_pages)
    if not doc_pages or doc_pages < 1:
        return txt, "last_page skipped (no doc_pages)"
    last_page = max(1, min(doc_pages, int(round(percent * doc_pages))))
    pattern = r'\["last_page"\]\s*=\s*\d+'
    replacement = f'["last_page"] = {last_page}'
    if re.search(pattern, txt):
        txt = re.sub(pattern, replacement, txt, count=1)
    else:
        txt = re.sub(r'(\}\s*)$', f'    {replacement},\n\\1', txt, count=1, flags=re.DOTALL)
    return txt, f"last_page->{last_page}"


def patch_lua(path, percent, status, today, device_doc_path, xpointer=None, fallback_pages=0):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        txt = f.read()
    changes = []
    txt, existed = patch_or_insert(txt, "percent_finished", f'["percent_finished"] = {percent:.14f}')
    changes.append("pct updated" if existed else "pct inserted")
    txt, existed = patch_or_insert(txt, "cre_dom_version", '["cre_dom_version"] = 20240114')
    changes.append("dom updated" if existed else "dom inserted")
    txt, existed = patch_or_insert(txt, "doc_path", f'["doc_path"] = "{device_doc_path}"')
    changes.append("doc_path updated" if existed else "doc_path inserted")
    lower = path.lower()
    if lower.endswith("metadata.epub.lua") and xpointer:
        txt = re.sub(r'\s*\["last_percent"\]\s*=\s*[^,\n]+,?\n', '\n', txt)
        txt, existed = patch_or_insert(txt, "last_xpointer", f'["last_xpointer"] = "{xpointer}"')
        changes.append("xpointer updated" if existed else "xpointer inserted")
    elif lower.endswith(("metadata.pdf.lua", "metadata.cbz.lua", "metadata.cbr.lua")):
        txt, change = patch_last_page(txt, percent, fallback_pages=fallback_pages)
        changes.append(change)

    m = re.search(r'\["summary"\]\s*=\s*\{(.*?)\}', txt, re.DOTALL)
    if m:
        inner = m.group(1)
        ni = inner
        if re.search(r'\["status"\]', ni):
            ni = re.sub(r'\["status"\]\s*=\s*"[^"]*"', f'["status"] = "{status}"', ni)
        else:
            ni += f'        ["status"] = "{status}",\n'
        if re.search(r'\["modified"\]', ni):
            ni = re.sub(r'\["modified"\]\s*=\s*"[^"]*"', f'["modified"] = "{today}"', ni)
        else:
            ni += f'        ["modified"] = "{today}",\n'
        txt = txt[:m.start()] + m.group(0).replace(m.group(1), ni) + txt[m.end():]
        changes.append("summary updated")
    else:
        blk = (f'    ["summary"] = {{\n'
               f'        ["modified"] = "{today}",\n'
               f'        ["status"] = "{status}",\n'
               f'    }},\n')
        txt = re.sub(r'(\}\s*)$', blk + r'\1', txt, count=1, flags=re.DOTALL)
        changes.append("summary inserted")

    with open(path, "w", encoding="utf-8") as f:
        f.write(txt)
    return changes


def create_lua(path, device_doc_path, percent, status, today, title, author, xpointer=None, doc_pages=0):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    t = title.replace('"', '\\"')
    a = author.replace('"', '\\"')
    lua_name = os.path.basename(path)
    lower = path.lower()
    lines = [
        f'-- {device_doc_path}.sdr/{lua_name}',
        'return {',
        '    ["annotations"] = {},',
        '    ["cre_dom_version"] = 20240114,',
        f'    ["doc_path"] = "{device_doc_path}",',
        '    ["doc_props"] = {',
        f'        ["authors"] = "{a}",',
        f'        ["title"] = "{t}",',
        '    },',
    ]
    if doc_pages and int(doc_pages) > 0:
        lines.append(f'    ["doc_pages"] = {int(doc_pages)},')
    lines.append(f'    ["percent_finished"] = {percent:.14f},')
    if lower.endswith("metadata.epub.lua") and xpointer:
        lines.append(f'    ["last_xpointer"] = "{xpointer}",')
    elif lower.endswith(("metadata.pdf.lua", "metadata.cbz.lua", "metadata.cbr.lua")) and doc_pages and int(doc_pages) > 0:
        last_page = max(1, min(int(doc_pages), int(round(percent * int(doc_pages)))))
        lines.append(f'    ["last_page"] = {last_page},')
    lines.extend([
        '    ["summary"] = {',
        f'        ["modified"] = "{today}",',
        f'        ["status"] = "{status}",',
        '    },',
        '}',
        ''
    ])
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))



def guess_page_count_from_file(book_path):
    lower = book_path.lower()
    try:
        if lower.endswith('.cbz'):
            with zipfile.ZipFile(book_path, 'r') as zf:
                exts = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.avif'}
                names = [n for n in zf.namelist() if os.path.splitext(n)[1].lower() in exts and not n.endswith('/')]
                return len(names) if names else 0
        if lower.endswith('.pdf'):
            from pypdf import PdfReader
            with open(book_path, 'rb') as f:
                reader = PdfReader(f)
                return len(reader.pages)
    except Exception:
        pass
    return 0

# ---------------- history ----------------

def parse_ts(s):
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return int(datetime.datetime.strptime(s[:19], fmt).timestamp())
        except ValueError:
            continue
    return int(datetime.datetime.now().timestamp())


def update_history(history_path, new_entries):
    existing = []
    seen = set()
    if os.path.exists(history_path):
        txt = open(history_path, encoding="utf-8", errors="replace").read()
        pairs = re.findall(r'\["time"\]\s*=\s*(\d+).*?\["file"\]\s*=\s*"([^"]+)"', txt, re.DOTALL)
        for ts_s, fp in pairs:
            if fp not in seen:
                existing.append((int(ts_s), fp))
                seen.add(fp)
    added = 0
    for dev_path, ts in new_entries:
        if dev_path not in seen:
            existing.append((ts, dev_path))
            seen.add(dev_path)
            added += 1
    existing.sort(key=lambda x: x[0], reverse=True)
    lines = ["return {\n"]
    for i, (ts, fp) in enumerate(existing, 1):
        lines.append(f'    [{i}] = {{\n        ["time"] = {ts},\n        ["file"] = "{fp}",\n    }},\n')
    lines.append("}\n")
    with open(history_path, "w", encoding="utf-8") as f:
        f.writelines(lines)
    return added

# ---------------- epub heuristic ----------------

def _read_xml_from_zip(zf, name):
    with zf.open(name) as f:
        return f.read()


def _find_opf_path(zf):
    try:
        raw = _read_xml_from_zip(zf, "META-INF/container.xml")
        root = ET.fromstring(raw)
        ns = {"c": "urn:oasis:names:tc:opendocument:xmlns:container"}
        rf = root.find(".//c:rootfile", ns)
        if rf is not None:
            return rf.attrib.get("full-path")
    except Exception:
        pass
    for n in zf.namelist():
        if n.lower().endswith(".opf"):
            return n
    return None


def _spine_hrefs(zf):
    opf_path = _find_opf_path(zf)
    if not opf_path:
        return None, [], {}
    raw = _read_xml_from_zip(zf, opf_path)
    root = ET.fromstring(raw)
    ns = {"opf": "http://www.idpf.org/2007/opf"}
    manifest = {}
    mani = root.find("opf:manifest", ns)
    if mani is not None:
        for item in mani:
            iid = item.attrib.get("id")
            href = item.attrib.get("href")
            if iid and href:
                manifest[iid] = href
    spine = []
    sp = root.find("opf:spine", ns)
    if sp is not None:
        for itemref in sp:
            iid = itemref.attrib.get("idref")
            href = manifest.get(iid)
            if href:
                spine.append(href)
    opf_dir = os.path.dirname(opf_path).replace("\\", "/")
    resolved, href_map = [], {}
    for href in spine:
        full = os.path.normpath(os.path.join(opf_dir, href)).replace("\\", "/")
        resolved.append(full)
        href_map[full.lower()] = len(resolved)
        href_map[os.path.basename(full).lower()] = len(resolved)
    return opf_path, resolved, href_map


def _strip_ns(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _first_chapter_xpointer(zf, chapter_path, fragment_index):
    try:
        raw = _read_xml_from_zip(zf, chapter_path)
        root = ET.fromstring(raw)
        body = None
        for elem in root.iter():
            if _strip_ns(elem.tag) == "body":
                body = elem
                break
        if body is None:
            return f"/body/DocFragment[{fragment_index}]/body"
        children = [c for c in list(body) if isinstance(c.tag, str)]
        container = children[0] if children and _strip_ns(children[0].tag) in ("section", "div", "article") else None
        if container is None:
            tags = [_strip_ns(c.tag) for c in children]
            h2_idx = [i+1 for i,t in enumerate(tags) if t == "h2"]
            if len(h2_idx) >= 2:
                return f"/body/DocFragment[{fragment_index}]/body/h2[2]/text()[1].0"
            if len(h2_idx) >= 1:
                return f"/body/DocFragment[{fragment_index}]/body/h2[1]/text()[1].0"
            for tag in ("h1", "h3", "h4", "p", "div"):
                idxs = [i+1 for i,t in enumerate(tags) if t == tag]
                if idxs:
                    return f"/body/DocFragment[{fragment_index}]/body/{tag}[{idxs[0]}]/text()[1].0"
            return f"/body/DocFragment[{fragment_index}]/body"
        ctag = _strip_ns(container.tag)
        inner = [c for c in list(container) if isinstance(c.tag, str)]
        tags = [_strip_ns(c.tag) for c in inner]
        h2_idx = [i+1 for i,t in enumerate(tags) if t == "h2"]
        if len(h2_idx) >= 2:
            return f"/body/DocFragment[{fragment_index}]/body/{ctag}/h2[2]/text()[1].0"
        if len(h2_idx) >= 1:
            return f"/body/DocFragment[{fragment_index}]/body/{ctag}/h2[1]/text()[1].0"
        for tag in ("h1", "h3", "h4", "p", "div"):
            idxs = [i+1 for i,t in enumerate(tags) if t == tag]
            if idxs:
                return f"/body/DocFragment[{fragment_index}]/body/{ctag}/{tag}[{idxs[0]}]/text()[1].0"
        return f"/body/DocFragment[{fragment_index}]/body/{ctag}"
    except Exception:
        return f"/body/DocFragment[{fragment_index}]/body"


def guess_epub_xpointer(book_path, chapter_file):
    if not chapter_file:
        return None, "no chapter file"
    if not (book_path.lower().endswith(".epub") or book_path.lower().endswith(".kepub.epub")):
        return None, "not epub-like"
    try:
        with zipfile.ZipFile(book_path, "r") as zf:
            _, spine, href_map = _spine_hrefs(zf)
            if not spine:
                return None, "no spine"
            key_full = chapter_file.replace("\\", "/").lower()
            key_base = os.path.basename(chapter_file).lower()
            frag = href_map.get(key_full) or href_map.get(key_base)
            if not frag:
                for href, idx in [(h, i+1) for i,h in enumerate(spine)]:
                    if os.path.basename(h).lower() == key_base:
                        frag = idx
                        key_full = href
                        break
            if not frag:
                return None, f"chapter not found in spine: {chapter_file}"
            actual = spine[frag-1]
            xp = _first_chapter_xpointer(zf, actual, frag)
            return xp, f"chapter-start heuristic from {actual} -> DocFragment[{frag}]"
    except Exception as e:
        return None, f"epub parse failed: {e}"

# ---------------- sandbox helpers ----------------

def write_text(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def copy_tree_filtered(src, dst, log):
    if not src or not os.path.isdir(src):
        log("WARN", "Starting with blank sandbox config")
        return
    if os.path.exists(dst):
        shutil.rmtree(dst, ignore_errors=True)
    def ignore(dirpath, names):
        ignored = set()
        if os.path.basename(dirpath) == "koreader":
            for n in names:
                if n == "plugins":
                    ignored.add(n)
        return ignored
    shutil.copytree(src, dst, ignore=ignore)
    log("INFO", f"Copying settings into sandbox:\n  {dst}")



def patch_settings_reader_lua(settings_path, sandbox_book_dir_wsl, log):
    ensure_dir(os.path.dirname(settings_path))
    if os.path.exists(settings_path):
        txt = open(settings_path, encoding="utf-8", errors="replace").read()
    else:
        txt = "return {\n}\n"
        log("WARN", "Starting with blank sandbox config")

    replacements = {
        'document_metadata_folder': '"doc"',
        'lastdir': f'"{sandbox_book_dir_wsl}"',
        'home_dir': f'"{sandbox_book_dir_wsl}"',
    }

    for key, value in replacements.items():
        pat = rf'\["{re.escape(key)}"\]\s*=\s*[^,\n]+'
        rep = f'["{key}"] = {value}'
        if re.search(pat, txt):
            txt = re.sub(pat, rep, txt)
        else:
            txt = re.sub(r'\}\s*$', f'    {rep},\n}}\n', txt, flags=re.DOTALL)

    write_text(settings_path, txt)
    log("INFO", f"Patched sandbox settings to force document_metadata_folder='doc':\n  {settings_path}")



def install_autogoto_plugin(sandbox_config_dir_win, target_book_wsl, target_percent, log):
    plugins_root = os.path.join(sandbox_config_dir_win, "plugins")
    plug_dir = os.path.join(plugins_root, "autogotopercent.koplugin")
    ensure_dir(plug_dir)

    meta = '''local _ = require("gettext")\nreturn {\n    name = "autogotopercent",\n    fullname = _("Auto Goto Percent"),\n    description = _("Automatically jumps to a configured percentage on book open."),\n}\n'''

    main = r'''local WidgetContainer = require("ui/widget/container/widgetcontainer")
local UIManager = require("ui/uimanager")
local Event = require("ui/event")
local logger = require("logger")
local Notification = require("ui/widget/notification")

local AutoGotoPercent = WidgetContainer:extend{
    name = "autogotopercent",
    is_doc_only = true,
}

local function read_target_file(path)
    local ok, data = pcall(dofile, path)
    if ok and type(data) == "table" then
        return data
    end
    return nil
end

function AutoGotoPercent:init()
    self._fired = false
end

function AutoGotoPercent:onReaderReady()
    if self._fired then
        return
    end

    local target_path = os.getenv("KO_AUTOGOTO_TARGET_FILE")
    if not target_path or target_path == "" then
        logger.info("AutoGotoPercent: KO_AUTOGOTO_TARGET_FILE not set")
        return
    end

    local target = read_target_file(target_path)
    if not target or not target.book or not target.percent then
        logger.info("AutoGotoPercent: invalid target file", target_path)
        return
    end

    local current = self.ui and self.ui.document and self.ui.document.file
    if current ~= target.book then
        logger.info("AutoGotoPercent: current book does not match target", current or "nil", target.book)
        return
    end

    self._fired = true
    logger.info("AutoGotoPercent: matched target book; scheduling goto", target.percent)

    local function do_jump(tag)
        if self.ui and self.ui.document then
            logger.info("AutoGotoPercent: sending GotoPercent event", tag, target.percent)
            self.ui:handleEvent(Event:new("GotoPercent", tonumber(target.percent)))
            UIManager:show(Notification:new{ text = string.format("Auto-jumped to %s%%", tostring(target.percent)) })
        end
    end

    local function close_and_quit()
        logger.info("AutoGotoPercent: closing reader and quitting KOReader")
        UIManager:broadcastEvent(Event:new("Close"))
        UIManager:nextTick(function()
            UIManager:quit(0)
        end)
    end

    -- Give KOReader a few chances to settle after opening and after first-run popups.
    UIManager:scheduleIn(1.0, function() do_jump("first") end)
    UIManager:scheduleIn(3.0, function() do_jump("second") end)
    UIManager:scheduleIn(7.0, close_and_quit)
end

return AutoGotoPercent
'''

    write_text(os.path.join(plug_dir, "_meta.lua"), meta)
    write_text(os.path.join(plug_dir, "main.lua"), main)

    target_cfg_dir = os.path.join(sandbox_config_dir_win, "autogoto")
    ensure_dir(target_cfg_dir)
    target_cfg_path_win = os.path.join(target_cfg_dir, "target.lua")
    target_cfg = (
        "return {\n"
        f"    book = {target_book_wsl!r},\n"
        f"    percent = {float(target_percent):.6f},\n"
        "}\n"
    )
    write_text(target_cfg_path_win, target_cfg)
    log("INFO", f"Installed sandbox AutoGoto plugin:\n  {plug_dir}")
    log("INFO", f"Wrote target config:\n  {target_cfg_path_win}")
    return target_cfg_path_win



def harvest_candidates(sandbox_home_win, out_dir, log):
    ensure_dir(out_dir)
    harvested = []
    for root, dirs, files in os.walk(sandbox_home_win):
        for f in files:
            if f.startswith("metadata.") and f.endswith(".lua"):
                src = os.path.join(root, f)
                rel = os.path.relpath(src, sandbox_home_win).replace("\\", "__")
                dst = os.path.join(out_dir, rel)
                ensure_dir(os.path.dirname(dst))
                shutil.copy2(src, dst)
                harvested.append(src)
                log("OK", f"Harvested sidecar candidate:\n  {src}")
                log("OK", f"Copied to:\n  {dst}")
    hist = os.path.join(sandbox_home_win, ".config", "koreader", "history.lua")
    if os.path.exists(hist):
        dst = os.path.join(out_dir, "history.lua")
        shutil.copy2(hist, dst)
        log("OK", f"Copied sandbox history.lua to {dst}")
    else:
        log("WARN", "Sandbox history.lua not found")
    if not harvested:
        log("WARN", "No changed/related sidecar found.")
    return harvested



def run_sandbox(appimage_win, book_win, settings_win, target_percent, workdir_win, out_dir_win,
                copy_settings, hide_console, log):
    if not os.path.exists(appimage_win):
        raise FileNotFoundError(appimage_win)
    if not os.path.exists(book_win):
        raise FileNotFoundError(book_win)

    if not workdir_win:
        workdir_win = os.path.join(os.path.expanduser("~"), "Desktop", "koreader_sandbox")
    if not out_dir_win:
        out_dir_win = os.path.join(workdir_win, "harvested")

    sandbox_home_win = os.path.join(workdir_win, "home")
    sandbox_books_win = os.path.join(sandbox_home_win, "books")
    sandbox_cfg_win = os.path.join(sandbox_home_win, ".config", "koreader")

    if os.path.exists(workdir_win):
        shutil.rmtree(workdir_win, ignore_errors=True)
    ensure_dir(sandbox_books_win)
    ensure_dir(out_dir_win)

    sandbox_book_win = os.path.join(sandbox_books_win, os.path.basename(book_win))
    shutil.copy2(book_win, sandbox_book_win)
    log("INFO", f"Copied book into sandbox:\n  {sandbox_book_win}")

    if copy_settings:
        copy_tree_filtered(settings_win, sandbox_cfg_win, log)
    else:
        ensure_dir(sandbox_cfg_win)
        log("WARN", "Starting with blank sandbox config")

    sandbox_books_wsl = win_to_wsl(sandbox_books_win)
    patch_settings_reader_lua(os.path.join(sandbox_cfg_win, "settings.reader.lua"), sandbox_books_wsl, log)

    sandbox_book_wsl = win_to_wsl(sandbox_book_win)
    target_cfg_win = install_autogoto_plugin(sandbox_cfg_win, sandbox_book_wsl, float(target_percent), log)
    target_cfg_wsl = win_to_wsl(target_cfg_win)

    appimage_wsl = win_to_wsl(appimage_win)
    sandbox_home_wsl = win_to_wsl(sandbox_home_win)

    log("INFO", f"Book: {book_win}\nSandbox copy: {sandbox_book_win}")
    log("INFO", f"Target percent: {target_percent}%")
    log("INFO", "AutoGoto plugin steps inside KOReader:")
    log("INFO", "  1) Dismiss any first-run popup if shown")
    log("INFO", "  2) Wait a few seconds for auto-jump")
    log("INFO", "  3) Close the book and exit KOReader")
    log("INFO", "")
    log("INFO", "Launching KOReader via WSL...")

    shell_cmd = (
        f'cd "{sandbox_home_wsl}" && '
        f'HOME="{sandbox_home_wsl}" '
        f'XDG_CONFIG_HOME="{sandbox_home_wsl}/.config" '
        f'XDG_DATA_HOME="{sandbox_home_wsl}/.local/share" '
        f'KO_AUTOGOTO_TARGET_FILE="{target_cfg_wsl}" '
        f'"{appimage_wsl}" --appimage-extract-and-run "{sandbox_book_wsl}"'
    )
    cmd = ["wsl", "bash", "-lc", shell_cmd]
    log("INFO", "WSL command prepared.")

    creationflags = 0
    if hide_console and hasattr(subprocess, "CREATE_NO_WINDOW"):
        creationflags = subprocess.CREATE_NO_WINDOW

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=creationflags,
    )
    stdout, stderr = proc.communicate()
    if stdout.strip():
        log("INFO", "WSL stdout:\n---------------------------------------------\n" + stdout)
    if stderr.strip():
        log("WARN", "WSL stderr:\n---------------------------------------------\n" + stderr)
    log("INFO", f"WSL exit code: {proc.returncode}")

    harvest_candidates(sandbox_home_win, out_dir_win, log)



def extract_last_xpointer(lua_path):
    try:
        txt = open(lua_path, encoding="utf-8", errors="replace").read()
        m = re.search(r'\["last_xpointer"\]\s*=\s*"([^"]+)"', txt)
        return m.group(1) if m else None
    except Exception:
        return None

def extract_percent_finished(lua_path):
    try:
        txt = open(lua_path, encoding="utf-8", errors="replace").read()
        m = re.search(r'\["percent_finished"\]\s*=\s*([0-9]+(?:\.[0-9]+)?)', txt)
        return float(m.group(1)) if m else None
    except Exception:
        return None


def generate_xpointer_via_sandbox(appimage_win, book_win, settings_win, percent, workdir_win, log_fn):
    out_dir_win = os.path.join(workdir_win, "harvested")
    run_sandbox(appimage_win, book_win, settings_win, percent * 100.0, workdir_win, out_dir_win, True, True, log_fn)
    # Prefer the live sandbox sidecar, fall back to harvested copies
    sandbox_book = os.path.join(workdir_win, "home", "books", os.path.basename(book_win))
    candidates = [get_sidecar_path(sandbox_book)]
    if os.path.isdir(out_dir_win):
        for root, dirs, files in os.walk(out_dir_win):
            for f in files:
                if f.startswith("metadata.") and f.endswith(".lua"):
                    candidates.append(os.path.join(root, f))
    for c in candidates:
        if os.path.exists(c):
            xp = extract_last_xpointer(c)
            if xp:
                log_fn("INFO", f"  sandbox xpointer: {xp}")
                return xp
    raise RuntimeError("Sandbox sidecar did not contain last_xpointer")

# ---------------- unified run ----------------

def run_unified(sqlite_path, books_folder, koreader_folder, dev_books, dry_run, existing_only,
                clear_cache, test_book_title, log_fn, done_fn, appimage_win, sandbox_workdir,
                manual_mode=False, ignore_if_koreader_ahead=False):
    today = datetime.date.today().strftime("%Y-%m-%d")
    counts = dict(matched=0, patched=0, created=0, history=0, skipped=0, errors=0, sandboxed=0)
    history_entries = []
    manual_output_root = get_manual_output_root() if manual_mode else None
    try:
        if manual_mode and not dry_run:
            if os.path.exists(manual_output_root):
                shutil.rmtree(manual_output_root, ignore_errors=True)
            os.makedirs(os.path.join(manual_output_root, "BOOKS"), exist_ok=True)
            os.makedirs(os.path.join(manual_output_root, "KOREADER"), exist_ok=True)
            log_fn("INFO", f"Manual Transfer Mode enabled. Output will be written to: {manual_output_root}")
        books = read_kobo_books(sqlite_path)
        log_fn("INFO", f"Found {len(books)} reading/completed books in Kobo database.")
        if test_book_title:
            books = [b for b in books if test_book_title.lower() in b["title"].lower()]
            log_fn("INFO", f"TEST MODE: filtered to {len(books)} book(s) matching '{test_book_title}'.")
        files = scan_books(books_folder)
        log_fn("INFO", f"Found {len(files)} book files in selected Books folder.")
        pairs = match_books(books, files)
        for idx, (book, book_path, reason) in enumerate(pairs, 1):
            if not book_path:
                counts["skipped"] += 1
                log_fn("SKIP", f"No match — {book['title']}")
                continue
            counts["matched"] += 1
            pct_display = f"{book['percent']*100:.0f}%"
            log_fn("MATCH", f"[{reason}] {book['title']} ({book['status']} {pct_display})")
            rel = os.path.relpath(book_path, books_folder).replace("\\", "/")
            dev_doc = (dev_books.rstrip("/") + "/" + rel) if dev_books else book_path.replace("\\", "/")
            log_fn("INFO", f"  device path: {dev_doc}")
            lua_path = get_manual_sidecar_path(book_path, books_folder) if manual_mode else get_sidecar_path(book_path)
            log_fn("INFO", f"  sidecar: {lua_path}")

            existing_sidecar_path = get_sidecar_path(book_path)
            if ignore_if_koreader_ahead and os.path.exists(existing_sidecar_path):
                existing_percent = extract_percent_finished(existing_sidecar_path)
                if existing_percent is not None and existing_percent > float(book["percent"]):
                    counts["skipped"] += 1
                    log_fn("SKIP", f"  KOReader ahead ({existing_percent*100:.1f}% > Kobo {book['percent']*100:.1f}%) — leaving existing sidecar unchanged")
                    continue

            history_entries.append((dev_doc, parse_ts(book["last_read"])))

            lower = book_path.lower()
            is_epub = lower.endswith(".epub") or lower.endswith(".kepub.epub")
            xpointer = None
            sandbox_needed = is_epub and book["status"] == "reading"
            if sandbox_needed:
                if dry_run:
                    log_fn("DRY", f"  would SANDBOX EPUB to generate xpointer at {pct_display}")
                else:
                    try:
                        xpointer = generate_xpointer_via_sandbox(appimage_win, book_path, koreader_folder,
                                                                 book["percent"], sandbox_workdir, log_fn)
                        counts["sandboxed"] += 1
                    except Exception as e:
                        counts["errors"] += 1
                        log_fn("ERROR", f"  sandbox failed: {e}")
                        continue
            elif is_epub:
                xpointer, why = guess_epub_xpointer(book_path, book["chapter_file"])
                if xpointer:
                    log_fn("INFO", f"  heuristic xpointer: {xpointer} [{why}]")
            # non-epub direct patch/create
            if dry_run:
                exists = os.path.exists(lua_path)
                action = "PATCH" if exists else "CREATE"
                log_fn("DRY", f"  would {action} (exists={exists})")
                continue
            try:
                removed = remove_old_metadata_backups(lua_path)
                for old in removed:
                    log_fn("INFO", f"  removed old sidecar backup: {old}")
                page_guess = book.get("num_pages", 0)
                if (not page_guess or int(page_guess) <= 0) and lower.endswith(('.cbz', '.pdf')):
                    page_guess = guess_page_count_from_file(book_path)
                    if page_guess:
                        src = 'CBZ archive' if lower.endswith('.cbz') else 'PDF file'
                        log_fn("INFO", f"  page count guessed from {src}: {page_guess}")
                if os.path.exists(lua_path):
                    changes = patch_lua(lua_path, book["percent"], book["status"], today, dev_doc,
                                        xpointer=xpointer, fallback_pages=page_guess)
                    counts["patched"] += 1
                    log_fn("PATCH", f"  → patched [{', '.join(changes)}]")
                elif existing_only:
                    counts["skipped"] += 1
                    log_fn("SKIP", "  no sidecar yet (existing-only mode)")
                else:
                    create_lua(lua_path, dev_doc, book["percent"], book["status"], today,
                               book["title"], book["author"], xpointer=xpointer, doc_pages=page_guess)
                    counts["created"] += 1
                    log_fn("CREATE", "  → created")
            except Exception as e:
                counts["errors"] += 1
                log_fn("ERROR", f"  write failed: {e}")
        if history_entries and not dry_run:
            if manual_mode:
                hist = os.path.join(manual_output_root, "KOREADER", "history.lua")
                n = update_history(hist, history_entries)
                counts["history"] = n
                log_fn("INFO", f"Prepared history.lua for manual copy (+{n} entries).")
            elif koreader_folder:
                hist = os.path.join(koreader_folder, "history.lua")
                if os.path.exists(hist):
                    n = update_history(hist, history_entries)
                    counts["history"] = n
                    log_fn("INFO", f"Updated history.lua (+{n} entries).")
        if clear_cache and koreader_folder and not dry_run and not manual_mode:
            cache = os.path.join(koreader_folder, "settings", "bookinfo_cache.sqlite3")
            if os.path.exists(cache):
                try:
                    os.remove(cache)
                    log_fn("INFO", "Deleted bookinfo cache; KOReader will rebuild it on next launch.")
                except Exception as e:
                    log_fn("ERROR", f"Could not delete cache: {e}")
    except Exception as e:
        counts["errors"] += 1
        import traceback
        log_fn("ERROR", f"Fatal: {e}")
        log_fn("ERROR", traceback.format_exc())
    if manual_mode and not dry_run:
        write_manual_instructions(dev_books, manual_output_root, log_fn)
    if os.path.exists(sandbox_workdir):
        try:
            shutil.rmtree(sandbox_workdir, ignore_errors=True)
            log_fn("INFO", f"Sandbox deleted after migration: {sandbox_workdir}")
        except Exception as e:
            log_fn("ERROR", f"Could not delete sandbox: {e}")
    done_fn(counts, dry_run)

# ---------------- GUI ----------------

def labeled_entry(parent, label_text, var, browse_cmd=None, hint=None):
    row = tk.Frame(parent, bg=PANEL)
    row.pack(fill="x", pady=3)
    tk.Label(row, text=label_text, bg=PANEL, fg=MUTED, font=FONT_SMALL, width=22, anchor="w").pack(side="left")
    entry = tk.Entry(row, textvariable=var, bg=ENTRY_BG, fg=TEXT, insertbackground=TEXT,
                     relief="flat", font=FONT_MONO, highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=ACCENT)
    entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
    if browse_cmd:
        tk.Button(row, text="Browse", command=browse_cmd, bg=PANEL2, fg=TEXT,
                  relief="flat", font=FONT_SMALL, padx=10, pady=3).pack(side="left")
    if hint:
        tk.Label(row, text=hint, bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(side="left", padx=6)
    return entry


def sep(parent):
    tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=8)


def sec(parent, text):
    tk.Label(parent, text=text, bg=PANEL, fg=MUTED, font=("Segoe UI", 8), anchor="w").pack(fill="x", pady=(8,2))


def chk(parent, var, label):
    f = tk.Frame(parent, bg=PANEL)
    f.pack(side="left", padx=(0, 20))
    tk.Checkbutton(f, variable=var, bg=PANEL, fg=TEXT, selectcolor=PANEL2,
                   activebackground=PANEL, activeforeground=TEXT,
                   relief="flat", font=FONT_SMALL).pack(side="left")
    tk.Label(f, text=label, bg=PANEL, fg=TEXT, font=FONT_SMALL).pack(side="left")


class App(tk.Tk):
    LOG_COLORS = {"INFO": MUTED, "MATCH": GREEN, "PATCH": ACCENT, "CREATE": AMBER,
                  "SKIP": MUTED, "DRY": AMBER, "ERROR": RED}
    def __init__(self):
        super().__init__()
        self.title("Kobo → KOReader Progress Migrator v1.0")
        self.configure(bg=BG)
        self.minsize(860, 760)
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=BG, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Kobo → KOReader", bg=BG, fg=TEXT, font=FONT_HEAD).pack(side="left")
        tk.Label(hdr, text="  Progress Migrator v1.0", bg=BG, fg=MUTED, font=FONT_UI).pack(side="left", pady=4)

        form = tk.Frame(self, bg=PANEL, padx=18, pady=14)
        form.pack(fill="x", padx=14, pady=(0, 6))
        sec(form, "SOURCE: Kobo Device")
        self.v_sqlite = tk.StringVar()
        labeled_entry(form, "KoboReader.sqlite", self.v_sqlite, lambda: self._pick_file(self.v_sqlite),
                      hint=r"<device>\.kobo\KoboReader.sqlite")
        sec(form, "TARGET: KOreader device")
        self.v_books = tk.StringVar()
        self.v_koreader = tk.StringVar()
        labeled_entry(form, "Books folder", self.v_books, lambda: self._pick_dir(self.v_books),
                      hint="Books folder on target KOreader device")
        labeled_entry(form, "KOReader folder", self.v_koreader, lambda: self._pick_dir(self.v_koreader),
                      hint=r"<device>\.adds\koreader  Will vary on Android Device")
        sep(form)
        sec(form, "TARGET DEVICE BASE PATH")
        tk.Button(form, text="⚙ Auto-detect from settings.reader.lua", command=self._auto_detect,
                  bg=PANEL2, fg=ACCENT, relief="flat", font=FONT_SMALL, padx=10, pady=4).pack(anchor="w", pady=(2,0))
        self.v_dev_books = tk.StringVar(value="/mnt/onboard")
        labeled_entry(form, "KOreader base path", self.v_dev_books,
                      hint="e.g Kobo: /mnt/onboard  Android: /storage/emulated/0/Books")
        sep(form)
        sec(form, "TEST MODE")
        self.v_test_title = tk.StringVar()
        labeled_entry(form, "Test book title", self.v_test_title,
                      hint="leave blank to process all books; recommended first")
        sep(form)
        sec(form, "OPTIONS")
        row = tk.Frame(form, bg=PANEL)
        row.pack(fill="x", pady=(2,0))
        self.v_dry = tk.BooleanVar(value=False)
        self.v_existing = tk.BooleanVar(value=False)
        self.v_cache = tk.BooleanVar(value=False)
        self.v_manual = tk.BooleanVar(value=False)
        self.v_ignore_ahead = tk.BooleanVar(value=False)
        chk(row, self.v_dry, "Dry run")
        chk(row, self.v_existing, "Patch existing only")
        chk(row, self.v_cache, "Delete bookinfo cache after run")
        chk(row, self.v_manual, "Manual Transfer Mode (Android / unstable connection)")
        chk(row, self.v_ignore_ahead, "If KOReader progress is ahead, ignore that book")

        tk.Label(form, text="Optional: use Manual Transfer Mode for Android Shared Internal Storage / CrossDevice paths that are unreliable or not mounted as a normal drive. The tool will build an OUTPUT folder next to the launcher with files ready to copy manually.", justify="left", wraplength=760, bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(fill="x", pady=(6,0))

        btn_row = tk.Frame(self, bg=BG, padx=14, pady=8)
        btn_row.pack(fill="x")
        self.btn_run = tk.Button(btn_row, text="▶  Run Migration", command=self._run,
                                 bg=ACCENT, fg="white", relief="flat", font=FONT_TITLE,
                                 padx=22, pady=10)
        self.btn_run.pack(side="left")
        self.lbl_status = tk.Label(btn_row, text="", bg=BG, fg=MUTED, font=FONT_SMALL)
        self.lbl_status.pack(side="left", padx=14)

        stats = tk.Frame(self, bg=BG, padx=14)
        stats.pack(fill="x")
        self._stat_vars = {}
        for key, label in [("matched","Matched"),("patched","Patched"),("created","Created"),
                           ("sandboxed","Sandboxed"),("history","History"),("skipped","Skipped"),("errors","Errors")]:
            v = tk.StringVar(value="—")
            self._stat_vars[key] = v
            card = tk.Frame(stats, bg=PANEL, padx=12, pady=8)
            card.pack(side="left", padx=(0, 6), pady=(0, 8))
            tk.Label(card, text=label, bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")
            color = RED if key == "errors" else (AMBER if key in ("skipped","sandboxed") else TEXT)
            tk.Label(card, textvariable=v, bg=PANEL, fg=color, font=("Segoe UI Semibold", 18)).pack(anchor="w")

        log_frame = tk.Frame(self, bg=BG)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        tk.Label(log_frame, text="LOG", bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")
        self.log = scrolledtext.ScrolledText(log_frame, bg="#111113", fg=MUTED,
                                             font=FONT_MONO, relief="flat", state="disabled", wrap="word",
                                             insertbackground=TEXT, highlightthickness=1,
                                             highlightbackground=BORDER, highlightcolor=BORDER)
        self.log.pack(fill="both", expand=True)
        for tag, color in self.LOG_COLORS.items():
            self.log.tag_config(tag, foreground=color)

    def _pick_file(self, var):
        p = filedialog.askopenfilename(title="Select file")
        if p:
            var.set(p)
    def _pick_dir(self, var):
        p = filedialog.askdirectory(title="Select folder")
        if p:
            var.set(p)
    def _log(self, level, msg):
        self.after(0, self._log_main, level, msg)
    def _log_main(self, level, msg):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{level:<8}] ", level)
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
    def _set_status(self, msg, color=MUTED):
        self.lbl_status.configure(text=msg, fg=color)

    def _auto_detect(self):
        kr = self.v_koreader.get().strip()
        books = self.v_books.get().strip()
        if not kr or not os.path.isdir(kr):
            messagebox.showinfo("Auto-detect", "Set the KOReader folder first.")
            return
        settings_path = find_settings_lua(kr)
        if not settings_path:
            messagebox.showinfo("Auto-detect", f"Could not find settings.reader.lua in:\n{kr}")
            return
        s = parse_settings_lua(settings_path)
        home = s.get("home_dir", "")
        if home:
            self.v_dev_books.set(derive_device_books_root(home, books))
        self._log("INFO", f"Auto-detected home_dir = '{home}'")
        self._log("INFO", f"Device books root = '{self.v_dev_books.get()}'")

    def _run(self):
        sqlite = self.v_sqlite.get().strip()
        books = self.v_books.get().strip()
        koreader = self.v_koreader.get().strip()
        dev_bks = self.v_dev_books.get().strip()
        dry = self.v_dry.get()
        existing = self.v_existing.get()
        cache = self.v_cache.get()
        manual = self.v_manual.get()
        ignore_ahead = self.v_ignore_ahead.get()
        test_title = self.v_test_title.get().strip()
        if not sqlite or not os.path.isfile(sqlite):
            messagebox.showerror("Error", "Select a valid KoboReader.sqlite file.")
            return
        if not books or not os.path.isdir(books):
            messagebox.showerror("Error", "Select a valid Books folder.")
            return
        if not koreader or not os.path.isdir(koreader):
            messagebox.showerror("Error", "Select a valid KOReader folder.")
            return
        if ":" in dev_bks or "\\" in dev_bks:
            messagebox.showerror("Error", "Device books root must be a Linux-style path like /mnt/onboard/Books.")
            return
        launcher_dir = os.path.dirname(os.path.abspath(__file__))
        appimage_win = os.path.join(launcher_dir, APPIMAGE_NAME)
        if not os.path.exists(appimage_win):
            messagebox.showerror("Error", f"Put {APPIMAGE_NAME} next to this launcher first.\nExpected:\n{appimage_win}")
            return
        sandbox_workdir = os.path.join(os.path.expanduser("~"), "Desktop", "koreader_sandbox_batch")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        for v in self._stat_vars.values():
            v.set("—")
        self.btn_run.configure(state="disabled", text="Running…")
        self._set_status("Running…", MUTED)
        self._log("INFO", f"Using AppImage next to launcher: {appimage_win}")

        def worker():
            run_unified(sqlite, books, koreader, dev_bks, dry, existing, cache, test_title,
                        self._log, self._on_done, appimage_win, sandbox_workdir,
                        manual_mode=manual, ignore_if_koreader_ahead=ignore_ahead)
        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, counts, dry_run):
        self.after(0, self._show_results, counts, dry_run)
    def _show_results(self, counts, dry_run):
        for key, v in self._stat_vars.items():
            v.set(str(counts.get(key, 0)))
        if dry_run:
            self._set_status(f"Dry run — {counts['matched']} books would be processed", AMBER)
        elif counts['errors']:
            self._set_status(f"Finished with {counts['errors']} error(s)", RED)
        else:
            self._set_status(f"Done — {counts['patched']} patched, {counts['created']} created, {counts['sandboxed']} sandboxed", GREEN)
        self.btn_run.configure(state="normal", text="▶  Run Migration")

if __name__ == "__main__":
    App().mainloop()
