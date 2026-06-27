"""
S/W要求仕様書(.docx)を、章番号パターン(例 3.2.1)で「項」単位に分割し、
人間レビュー用の導線として以下を出力する:

  out/
   ├─ index.html          … 目次(章階層ネスト, 各項へジャンプ)
   ├─ index.md            … 目次(Markdown)
   ├─ md/<番号>_<題>.md    … 項ごとのMarkdown
   └─ html/<番号>_<題>.html… 項ごとのHTML(要求IDにアンカー付与)

設計方針:
- 見出しは「段落スタイル名に Heading/見出し を含む」または
  「テキスト先頭が章番号パターン」で判定（スタイルが付いていない実物にも対応）。
- 章番号(1, 1.1, 2.1.3 ...)を見出しテキスト先頭から抽出し、階層と分割境界を決める。
- 分割の粒度は「見つかった最も深い番号付き見出し＝項」単位。
  ある番号付き見出しから次の番号付き見出しの直前までを1つの項とする。
- 表は Markdown表 / HTML表 に変換して各項に含める。
- LLMによる要約や改変は一切しない（原本の忠実な分割のみ）。

使い方:
  python convert.py input.docx out/
"""
import os
import re
import sys
import html as html_lib
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

# 章番号パターン: 先頭の "1" "1.2" "10.3.4" など。後ろに空白か全角空白か文末。
NUM_RE = re.compile(r'^\s*(\d+(?:\.\d+)*)[\s　]+(.*\S)?\s*$')
# 要求IDパターン（必要に応じて実物に合わせて調整）
REQID_RE = re.compile(r'\b(REQ-\d+)\b')


def iter_block_items(parent):
    """document本文を、段落と表が現れた順に返す。"""
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)


def is_heading_style(para):
    name = (para.style.name or "") if para.style else ""
    return ("Heading" in name) or ("見出し" in name) or (name == "Title")


def parse_heading(para):
    """段落が「番号付き見出し」なら (番号, 題名, 階層レベル) を返す。違えば None。
    見出しスタイルが付いていれば優先採用。番号が取れる場合は番号で階層を決める。
    """
    text = para.text.strip()
    if not text:
        return None
    m = NUM_RE.match(text)
    styled = is_heading_style(para)
    if m:
        num = m.group(1)
        title = (m.group(2) or "").strip()
        level = num.count(".") + 1  # 1 -> L1, 1.2 -> L2, 2.1.3 -> L3
        return (num, title, level)
    # 番号は無いがスタイル上は見出し（例: 「概要」だけの見出し）→ 番号なし見出し扱い
    if styled:
        return ("", text, 1)
    return None


def table_to_md(table):
    rows = []
    for r in table.rows:
        cells = [c.text.replace("\n", " ").strip() for c in r.cells]
        rows.append(cells)
    if not rows:
        return ""
    out = []
    header = rows[0]
    out.append("| " + " | ".join(header) + " |")
    out.append("| " + " | ".join("---" for _ in header) + " |")
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return "\n".join(out)


def table_to_html(table):
    out = ['<table>']
    for i, r in enumerate(table.rows):
        out.append("<tr>")
        tag = "th" if i == 0 else "td"
        for c in r.cells:
            cell = html_lib.escape(c.text.strip()).replace("\n", "<br>")
            out.append(f"<{tag}>{cell}</{tag}>")
        out.append("</tr>")
    out.append("</table>")
    return "\n".join(out)


def slugify(num, title):
    base = (num + "_" + title) if num else title
    base = base.strip().replace(" ", "_").replace("　", "_")
    base = re.sub(r'[^\w\.\-一-龠ぁ-んァ-ヶー]', "", base)
    return base or "section"


def add_reqid_anchors_html(text_html):
    """HTML本文中の REQ-xxxx に id とアンカーを付ける（検索ジャンプ用）。"""
    def repl(m):
        rid = m.group(1)
        return f'<span class="reqid" id="{rid}">{rid}</span>'
    return REQID_RE.sub(repl, text_html)


def convert(docx_path, out_dir):
    doc = Document(docx_path)
    blocks = list(iter_block_items(doc))

    # まず番号付き(またはスタイル)見出しの位置を集める
    sections = []  # 各要素: {num,title,level, blocks:[...]}
    current = None
    preamble = []  # 最初の見出しより前（表題など）

    for blk in blocks:
        if isinstance(blk, Paragraph):
            h = parse_heading(blk)
            if h is not None:
                num, title, level = h
                current = {"num": num, "title": title, "level": level, "blocks": []}
                sections.append(current)
                continue
        if current is None:
            preamble.append(blk)
        else:
            current["blocks"].append(blk)

    os.makedirs(os.path.join(out_dir, "md"), exist_ok=True)
    os.makedirs(os.path.join(out_dir, "html"), exist_ok=True)

    toc_entries = []  # (num, title, level, md_name, html_name)

    for sec in sections:
        num, title, level = sec["num"], sec["title"], sec["level"]
        slug = slugify(num, title)
        md_name = f"{slug}.md"
        html_name = f"{slug}.html"

        # --- Markdown本文 ---
        md_lines = []
        heading_prefix = "#" * min(level + 1, 6)
        head_text = (f"{num} {title}" if num else title).strip()
        md_lines.append(f"{heading_prefix} {head_text}\n")
        for blk in sec["blocks"]:
            if isinstance(blk, Paragraph):
                t = blk.text.strip()
                if t:
                    md_lines.append(t + "\n")
            elif isinstance(blk, Table):
                md_lines.append(table_to_md(blk) + "\n")
        md_body = "\n".join(md_lines).strip() + "\n"
        with open(os.path.join(out_dir, "md", md_name), "w", encoding="utf-8") as f:
            f.write(md_body)

        # --- HTML本文 ---
        html_parts = []
        for blk in sec["blocks"]:
            if isinstance(blk, Paragraph):
                t = blk.text.strip()
                if t:
                    esc = add_reqid_anchors_html(html_lib.escape(t))
                    html_parts.append(f"<p>{esc}</p>")
            elif isinstance(blk, Table):
                html_parts.append(table_to_html(blk))
        sec_id = num if num else slug
        html_doc = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8">
<title>{html_lib.escape(head_text)}</title>
<style>
 body{{font-family:system-ui,"Hiragino Sans","Noto Sans JP",sans-serif;max-width:820px;margin:2rem auto;padding:0 1rem;line-height:1.7;color:#1a1a1a}}
 h1{{font-size:1.4rem;border-bottom:2px solid #2E75B6;padding-bottom:.3rem}}
 .reqid{{background:#fff3cd;font-weight:600;padding:0 .2em;border-radius:3px}}
 table{{border-collapse:collapse;margin:1rem 0;width:100%}}
 th,td{{border:1px solid #ccc;padding:.4rem .6rem;text-align:left}}
 th{{background:#D5E8F0}}
 nav a{{color:#2E75B6}}
</style></head>
<body>
<nav><a href="../index.html">← 目次へ戻る</a></nav>
<h1 id="{html_lib.escape(sec_id)}">{html_lib.escape(head_text)}</h1>
{chr(10).join(html_parts)}
</body></html>
"""
        with open(os.path.join(out_dir, "html", html_name), "w", encoding="utf-8") as f:
            f.write(html_doc)

        toc_entries.append((num, title, level, md_name, html_name))

    # --- 目次(Markdown) ---
    md_toc = ["# 目次\n"]
    for num, title, level, md_name, html_name in toc_entries:
        indent = "  " * (level - 1)
        label = (f"{num} {title}" if num else title).strip()
        md_toc.append(f"{indent}- [{label}](md/{md_name})")
    with open(os.path.join(out_dir, "index.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md_toc) + "\n")

    # --- 目次(HTML, 章番号インクリメンタル絞り込み付き) ---
    rows_html = []
    for num, title, level, md_name, html_name in toc_entries:
        label = (f"{num} {title}" if num else title).strip()
        rows_html.append(
            f'<li class="lv{level}" data-key="{html_lib.escape(label)}">'
            f'<a href="html/{html_name}">{html_lib.escape(label)}</a></li>'
        )
    index_html = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8">
<title>目次 — 要求仕様書</title>
<style>
 body{{font-family:system-ui,"Hiragino Sans","Noto Sans JP",sans-serif;max-width:820px;margin:2rem auto;padding:0 1rem;color:#1a1a1a}}
 h1{{font-size:1.5rem;border-bottom:2px solid #2E75B6;padding-bottom:.3rem}}
 #q{{width:100%;padding:.5rem;font-size:1rem;margin:1rem 0;box-sizing:border-box;border:1px solid #aaa;border-radius:4px}}
 ul{{list-style:none;padding-left:0}}
 li{{margin:.15rem 0}}
 li.lv1{{margin-top:.5rem;font-weight:600}}
 li.lv2{{padding-left:1.2rem;font-weight:500}}
 li.lv3{{padding-left:2.4rem;font-weight:400}}
 li.lv4{{padding-left:3.6rem}}
 a{{color:#2E75B6;text-decoration:none}} a:hover{{text-decoration:underline}}
 .hidden{{display:none}}
</style></head>
<body>
<h1>要求仕様書 — 目次</h1>
<input id="q" type="text" placeholder="章番号・タイトル・要求IDで絞り込み（例: 2.1 / 応答 / REQ-1001）">
<ul id="toc">
{chr(10).join(rows_html)}
</ul>
<script>
const q=document.getElementById('q'),items=[...document.querySelectorAll('#toc li')];
q.addEventListener('input',()=>{{
  const v=q.value.trim().toLowerCase();
  items.forEach(li=>{{
    const k=li.dataset.key.toLowerCase();
    li.classList.toggle('hidden', v && !k.includes(v));
  }});
}});
</script>
</body></html>
"""
    with open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)

    return len(toc_entries)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("usage: python convert.py input.docx out_dir/")
        sys.exit(1)
    n = convert(sys.argv[1], sys.argv[2])
    print(f"done: {n} sections -> {sys.argv[2]}")
