#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""docx -> 静态站（stdlib only，不依赖 python-docx / lxml）

用法: python docx2site.py <input.docx> <outdir>
产出: <outdir>/index.html  <outdir>/assets/style.css  <outdir>/assets/*.png
"""
import sys, os, re, zipfile, shutil, html
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(tag, ns=W):
    return "{%s}%s" % (ns, tag)


def text_of(el):
    out = []
    for node in el.iter():
        if node.tag == q("t"):
            out.append(node.text or "")
        elif node.tag == q("br"):
            out.append("\n")
        elif node.tag == q("tab"):
            out.append("\t")
    return "".join(out)


def run_is_bold(r):
    rPr = r.find(q("rPr"))
    if rPr is None:
        return False
    b = rPr.find(q("b"))
    return b is not None and b.get(q("val"), "true") not in ("0", "false")


def inline_html(p, rels, media_out, img_counter, out_dir=None):
    """段落内容 -> html（保留粗体 / 图片）"""
    parts = []
    for r in p.iter(q("r")):
        # 图片
        drew = False
        for blip in r.iter(q("blip", A)):
            rid = blip.get(q("embed", R))
            target = rels.get(rid)
            if target:
                name = os.path.basename(target)
                src = os.path.join(media_out, name)
                dst_dir = out_dir or media_out
                if os.path.exists(src):
                    fname = "%02d-%s" % (img_counter[0], name)
                    shutil.copy(src, os.path.join(dst_dir, fname))
                    img_counter[0] += 1
                    parts.append('<figure><img src="assets/%s" alt="图表"></figure>' % fname)
                    drew = True
        if drew:
            continue
        t = ""
        for node in r.iter():
            if node.tag == q("t"):
                t += node.text or ""
            elif node.tag == q("br"):
                t += "\n"
        if not t:
            continue
        t = html.escape(t)
        t = t.replace("\n", "<br>")
        if run_is_bold(r):
            t = "<strong>%s</strong>" % t
        parts.append(t)
    return "".join(parts)


def cell_html(tc, rels, media_out, img_counter, out_dir=None):
    spans = []
    tcPr = tc.find(q("tcPr"))
    colspan = 1
    if tcPr is not None:
        gs = tcPr.find(q("gridSpan"))
        if gs is not None:
            try:
                colspan = int(gs.get(q("val")))
            except Exception:
                colspan = 1
    for p in tc.findall(q("p")):
        spans.append(inline_html(p, rels, media_out, img_counter, out_dir))
    txt = "<br>".join(s for s in spans if s.strip())
    return '<td colspan="%d">%s</td>' % (colspan, txt)


def table_html(tbl, rels, media_out, img_counter, out_dir=None):
    rows = []
    for tr in tbl.findall(q("tr")):
        cells = [cell_html(tc, rels, media_out, img_counter, out_dir) for tc in tr.findall(q("tc"))]
        rows.append("<tr>%s</tr>" % "".join(cells))
    return '<div class="table-wrap"><table>%s</table></div>' % "".join(rows)


def main(src, outdir):
    z = zipfile.ZipFile(src)
    rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
    rels = {}
    for rel in ET.fromstring(rels_xml):
        rels[rel.get("Id")] = rel.get("Target")

    media_out = os.path.join(outdir, "assets")
    os.makedirs(media_out, exist_ok=True)
    tmp_media = os.path.join(outdir, "assets", "_raw")
    os.makedirs(tmp_media, exist_ok=True)
    for n in z.namelist():
        if n.startswith("word/media/"):
            with open(os.path.join(tmp_media, os.path.basename(n)), "wb") as f:
                f.write(z.read(n))

    root = ET.fromstring(z.read("word/document.xml").decode("utf-8"))
    body = root.find(q("body"))
    img_counter = [1]
    out = []
    toc = []
    hid = 0
    title_done = False

    for el in body:
        if el.tag == q("p"):
            pPr = el.find(q("pPr"))
            style = ""
            centered = False
            if pPr is not None:
                ps = pPr.find(q("pStyle"))
                if ps is not None:
                    style = ps.get(q("val"), "")
                jc = pPr.find(q("jc"))
                if jc is not None and jc.get(q("val")) == "center":
                    centered = True
            inner = inline_html(el, rels, tmp_media, img_counter, media_out)
            plain = text_of(el).strip()
            if not inner.strip() and not plain:
                continue
            if style.lower().startswith("heading"):
                n = int(re.sub(r"\D", "", style) or 1)
                lvl = str(min(n + 1, 6))
                hid += 1
                anchor = "h%d" % hid
                cls = ' class="center"' if centered else ""
                out.append('<h%s id="%s"%s>%s</h%s>' % (lvl, anchor, cls, inner, lvl))
                if lvl in ("2", "3"):
                    toc.append((lvl, anchor, plain))
            elif not title_done and len(plain) < 40 and inner:
                out.append('<p class="doc-title">%s</p>' % inner)
                title_done = True
            elif re.match(r"^(图|表)\s*\d", plain) and len(plain) < 120:
                out.append('<p class="caption">%s</p>' % inner)
            else:
                cls = ' class="center"' if centered else ""
                out.append("<p%s>%s</p>" % (cls, inner))
        elif el.tag == q("tbl"):
            out.append(table_html(el, rels, tmp_media, img_counter, media_out))

    # 目录
    toc_html = ['<nav class="toc"><p class="toc-h">目录</p><ul>']
    for lvl, anchor, txt in toc:
        toc_html.append('<li class="lvl%s"><a href="#%s">%s</a></li>' % (lvl, anchor, html.escape(txt)))
    toc_html.append("</ul></nav>")

    # 找出正文里第一个 h2 之前的内容（封面/阅读说明）
    try:
        first_h2 = next(i for i, s in enumerate(out) if s.startswith("<h2"))
    except StopIteration:
        first_h2 = 0
    head_part = out[:first_h2]
    rest = out[first_h2:]

    pdf_path = os.path.join(outdir, "assets", "report.pdf")
    pdf_link = ('<p class="dl"><a href="assets/report.pdf">下载 PDF 版（约 1 MB）</a></p>'
                if os.path.exists(pdf_path) else "")
    page = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>中国人口结构全景分层报告（区域细化版）</title>
<meta name="description" content="把 14 亿中国人切成十一条年龄带与七层身份：每一层的人数、钱和来源。数据基准 2025 年末。">
<link rel="stylesheet" href="assets/style.css">
</head>
<body>
<main class="wrap">
<header class="cover">
%s
%s
</header>
%s
<article class="doc">
%s
</article>
<footer class="site-foot">
<p>本页为静态存档版。引用请以文中标注的数据源为准，区分「官方公布」「据……推算」「作者估算」三类数字。</p>
</footer>
</main>
</body>
</html>
""" % ("\n".join(head_part), pdf_link, "\n".join(toc_html), "\n".join(rest))

    with open(os.path.join(outdir, "index.html"), "w", encoding="utf-8") as f:
        f.write(page)
    shutil.rmtree(tmp_media, ignore_errors=True)
    print("head blocks:", len(head_part), "body blocks:", len(rest), "toc:", len(toc), "images:", img_counter[0] - 1)
    print("html bytes:", len(page))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
