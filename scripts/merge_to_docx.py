#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_to_docx.py v2 — 论文模板内容整合器

按 Word 模板的分节结构工作：
  1. 摘要 / Abstract / 结论 / 参考文献 / 致谢：直接替换模板占位文字（保留标题与页眉页脚）
  2. 正文章节：每章新建分节插入目录之后，页眉自动显示本章标题（克隆模板页眉样式，横线一致）
  3. 附录：参考文献之后新建分节，页眉"附  录"
  4. 自动设置 updateFields，Word 打开时提示更新目录域

用法:
  python3 merge_to_docx.py --template 模板.docx --output 成品.docx \
      --abstract 摘要.txt --conclusion 结论.txt --references 参考文献.txt \
      --appendix 附录.txt:1 \
      第1章.txt 第2章.txt 第3章.txt ...

依赖: python-docx (pip3 install python-docx)
"""

import argparse
import copy
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

IMG_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff'}
TABLE_EXTS = {'.csv'}
TEXT_EXTS = {'.txt', '.md', '.markdown'}

# ------------------------- 基础工具 -------------------------

def par_text_of(elm):
    """XML 元素内所有 w:t 拼接"""
    return ''.join(t.text or '' for t in elm.iter(qn('w:t')))

def norm(s):
    """去掉所有空白，用于标题比对"""
    return re.sub(r'\s+', '', s or '')

def find_style(doc, candidates):
    for name in candidates:
        try:
            return doc.styles[name]
        except KeyError:
            continue
    return None

def style_of(doc, kind):
    table = {
        'h1': ['heading 1', 'Heading 1', '标题 1'],
        'h2': ['heading 2', 'Heading 2', '标题 2'],
        'h3': ['heading 3', 'Heading 3', '标题 3'],
        'body': ['Normal', '正文'],
        'caption': ['caption', 'Caption', '题注'],
    }
    return find_style(doc, table.get(kind, []))

# ------------------------- 公式转换（LaTeX → Word OMML） -------------------------

PANDOC = shutil.which('pandoc')

def _pandoc_parse(latex, block):
    """用 pandoc 把 LaTeX 公式转成 OMML。返回 lxml 段落根元素；失败返回 None"""
    if not PANDOC:
        return None
    md = f'$${latex}$$' if block else f'${latex}$'
    with tempfile.TemporaryDirectory() as td:
        md_path = os.path.join(td, 'f.md')
        out_path = os.path.join(td, 'f.docx')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md)
        r = subprocess.run([PANDOC, '-f', 'markdown', '-t', 'docx',
                            '-o', out_path, md_path], capture_output=True)
        if r.returncode != 0:
            return None
        try:
            with zipfile.ZipFile(out_path) as z:
                xml = z.read('word/document.xml')
        except (KeyError, OSError):
            return None
    try:
        from lxml import etree
        root = etree.fromstring(xml)
    except Exception:
        return None
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
    ps = root.findall('.//w:p', ns)
    return ps[0] if ps else None

def omml_block_element(par_root):
    """块公式：取段落内 m:oMathPara 元素"""
    for ch in par_root:
        if ch.tag == qn('m:oMathPara'):
            return ch
    return None

def omml_inline_children(par_root):
    """行内公式：按顺序取段落内的 w:r 文本与 m:oMath 子元素"""
    out = []
    for ch in par_root:
        if ch.tag in (qn('w:r'), qn('m:oMath')):
            out.append(ch)
    return out

# ------------------------- 三线表边框 -------------------------

def apply_three_line_borders(tbl):
    """三线表：顶线/底线 1.5磅(sz=12)，表头下线 0.75磅(sz=6)，无竖线与内部横线；表格整体居中"""
    tblPr = tbl._tbl.tblPr
    for el in tblPr.findall(qn('w:jc')):
        tblPr.remove(el)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is not None:
        tblW.addnext(jc)     # jc 须位于 tblW 之后，符合 tblPr 元素顺序
    else:
        tblPr.insert(0, jc)
    for el in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(el)
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        if edge in ('top', 'bottom'):
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '12')
        else:
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        tblBorders.append(el)
    tblPr.append(tblBorders)
    # 表头行底线（细线）
    if len(tbl.rows) >= 1:
        for cell in tbl.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn('w:tcBorders')):
                tcPr.remove(el)
            tcBorders = OxmlElement('w:tcBorders')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), '000000')
            tcBorders.append(bottom)
            tcPr.append(tcBorders)

# ------------------------- 字体设置 -------------------------

def set_run_fonts(run, east='宋体'):
    """数字/英文用 Times New Roman（ascii/hAnsi），中文用 east 指定字体（默认宋体）"""
    run.font.name = 'Times New Roman'
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), east)

# ------------------------- 标题识别 -------------------------

MD_HEADING = re.compile(r'^(#{1,4})\s+(.*)$')
NUM_H1 = re.compile(r'^(\d+)?\s*第[一二三四五六七八九十百\d]+章\s*(.*)$')
NUM_H2 = re.compile(r'^(\d{1,2})\.(\d{1,2})\s*(\D.*)$')
NUM_H3 = re.compile(r'^(\d{1,2})\.(\d{1,2})\.(\d{1,2})\s*(\D.*)$')
MD_TABLE_ROW = re.compile(r'^\s*\|(.+)\|\s*$')
MD_TABLE_SEP = re.compile(r'^\s*\|[\s:\-|]+\|\s*$')
MD_IMAGE = re.compile(r'^!\[(.*?)\]\((.+?)\)\s*$')

def classify_heading(line):
    """识别标题层级，返回 (level, text) 或 None"""
    m = MD_HEADING.match(line)
    if m:
        return min(len(m.group(1)), 3), m.group(2).strip()
    s = line.strip()
    if NUM_H3.match(s):
        return 3, s
    if NUM_H2.match(s):
        return 2, s
    if NUM_H1.match(s):
        return 1, s
    return None

def parse_table_rows(lines, i):
    rows = []
    while i < len(lines):
        line = lines[i]
        if not MD_TABLE_ROW.match(line):
            break
        if MD_TABLE_SEP.match(line):
            i += 1
            continue
        rows.append([c.strip() for c in line.strip().strip('|').split('|')])
        i += 1
    return rows, i

def read_lines(path):
    with open(path, 'r', encoding='utf-8-sig') as f:
        return f.read().split('\n')

# ------------------------- 写入器（XML级光标插入） -------------------------

class CursorWriter:
    """在锚点元素之前（replace 模式）或光标之后（chapter 模式）写入内容"""

    def __init__(self, doc, mode='before', anchor=None):
        self.doc = doc
        self.mode = mode          # 'before': 插到 anchor 前；'after': 光标依次后移
        self.anchor = anchor      # 'before' 模式的锚点元素
        self.cursor = anchor      # 'after' 模式的光标
        self.fig_no = 0
        self.tab_no = 0
        self.last_elm = None
        self.toc_skip_headings = False
        self._bm_next = self._max_bookmark_id(doc) + 1   # 避开模板已有书签 id

    @staticmethod
    def _max_bookmark_id(doc):
        """文档中现有 bookmarkStart 的最大 id（模板 TOC 书签可能占用 1..n）"""
        ids = []
        for el in doc.element.body.iter(qn('w:bookmarkStart')):
            v = el.get(qn('w:id'))
            if v and v.isdigit():
                ids.append(int(v))
        return max(ids) if ids else 0

    def _place(self, elm):
        if self.mode == 'before':
            self.anchor.addprevious(elm)
        else:
            self.cursor.addnext(elm)
            self.cursor = elm
        self.last_elm = elm
        return elm

    def _new_par(self, style_kind=None):
        p = OxmlElement('w:p')
        style = style_of(self.doc, style_kind) if style_kind else None
        if style is not None:
            pPr = OxmlElement('w:pPr')
            ps = OxmlElement('w:pStyle')
            ps.set(qn('w:val'), style.style_id)
            pPr.append(ps)
            p.append(pPr)
        self._place(p)
        return Paragraph(p, self.doc._body)

    # --- 内容方法 ---
    def heading(self, level, text, toc_skip=False):
        par = self._new_par(f'h{level}')
        pPr = par._p.get_or_add_pPr()
        if level == 3:
            # 三级/四级标题：段前保持(10.25磅)、段后 0 与正文紧贴，行距 1.25 倍(line=300 auto)
            # （heading 3 样式自带的 line=324 auto + 段后205 使标题与正文间距偏大）
            sp = OxmlElement('w:spacing')
            sp.set(qn('w:before'), '205')
            sp.set(qn('w:after'), '0')
            sp.set(qn('w:line'), '300')
            sp.set(qn('w:lineRule'), 'auto')
            pPr.append(sp)
        if toc_skip:
            # outlineLvl=9 = 正文级别，TOC(\o "1-3") 不会收录
            lvl = OxmlElement('w:outlineLvl')
            lvl.set(qn('w:val'), '9')
            pPr.append(lvl)
        par.add_run(text)
        return par

    def _apply_body_fmt(self, par):
        """与模板正文一致的局部格式：20磅固定行距(line=400 exact) + 首行缩进2字符
        （spacing 需在 ind 之前，符合 pPr 元素顺序）"""
        pPr = par._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:line'), '400')
        sp.set(qn('w:lineRule'), 'exact')
        pPr.append(sp)
        ind = OxmlElement('w:ind')
        ind.set(qn('w:firstLine'), '420')
        ind.set(qn('w:firstLineChars'), '200')
        pPr.append(ind)

    def body(self, text, sup_refs=True):
        par = self._new_par('body')
        self._apply_body_fmt(par)
        self._write_rich(par, text, sup_refs=sup_refs)
        return par

    def reference(self, text):
        """参考文献条目：与模板一致的格式 —— 12磅(sz=24) + 1.25倍行距(line=300 auto)
        + 悬挂缩进；编号 `[n]` 后两空格，悬挂缩进=编号宽+两空格宽（按位数动态），
        换行行与首行文献内容对齐且编号后间距统一适中（模板 369=1位数字+两空格）"""
        par = self._new_par('body')
        pPr = par._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:line'), '300')
        sp.set(qn('w:lineRule'), 'auto')
        pPr.append(sp)
        m = re.match(r'^\[(\d+)\](.*)$', text, re.S)
        if m:
            num, rest = m.group(1), m.group(2)
            # 编号宽：方括号70twips/个 + 数字130twips/个（TNR 12磅）；两空格=240twips
            hang = 140 + 130 * len(num) + 240
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), str(hang))
            ind.set(qn('w:hanging'), str(hang))
            pPr.append(ind)
            self._add_run(par, f'[{num}]')          # 编号 [n]
            self._add_run(par, '  ')                 # 编号与文献间距（同模板两空格）
            if rest:
                self._add_run(par, rest)             # 文献内容
        else:
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), '510')
            ind.set(qn('w:hanging'), '510')
            pPr.append(ind)
            self._add_run(par, text)
        for r in par.runs:
            rPr = r._r.get_or_add_rPr()
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '24')
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), '24')
            rPr.append(szCs)
        return par

    def caption(self, text):
        """居中表注/图注（无首行缩进）"""
        par = self._new_par()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._write_rich(par, text)
        return par

    def _add_run(self, par, text, bold=False, sup=False):
        """创建 run 并统一字体：数字/英文 Times New Roman，中文宋体"""
        run = par.add_run(text)
        if bold:
            run.bold = True
        if sup:
            run.font.superscript = True
        set_run_fonts(run)
        return run

    def _add_fld(self, par, fldtype, text='', sup=False):
        """构造域 run（HYPERLINK 域用）：fldtype ∈ begin/instr/separate/end。
        HYPERLINK 域是 Word/WPS 内部跳转（Ctrl+点击）最兼容的方式"""
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rf = OxmlElement('w:rFonts')
        rf.set(qn('w:ascii'), 'Times New Roman')
        rf.set(qn('w:hAnsi'), 'Times New Roman')
        rf.set(qn('w:eastAsia'), '宋体')
        rPr.append(rf)
        if sup:
            va = OxmlElement('w:vertAlign')
            va.set(qn('w:val'), 'superscript')
            rPr.append(va)
        r.append(rPr)
        if fldtype == 'instr':
            it = OxmlElement('w:instrText')
            it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            it.text = text
            r.append(it)
        else:
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), fldtype)
            r.append(fc)
        par._p.append(r)
        return r

    # 引用标记：[1]、[1-3]、[1,2]、[5][6] 等纯数字引用
    _CITE_RE = r'\[\d+(?:[-,，、]\s*\d+)*\]'

    def _emit_refs(self, par, text, sup_refs):
        """文本中 [n] 文献引用 → 上标+HYPERLINK 域（Ctrl+点击跳转到参考文献书签）；
        sup_refs=False 时整体原样输出"""
        if not sup_refs:
            self._add_run(par, text)
            return
        for piece in re.split(f'({self._CITE_RE})', text):
            if re.match(f'^{self._CITE_RE}$', piece):
                self._emit_cite(par, piece)
            elif piece:
                self._add_run(par, piece)

    def _emit_cite(self, par, piece):
        """把引用片段（[1]、[1-3]、[1,2]、[5][6]）写成上标 HYPERLINK 域。
        连续引用 [5][6] 拆开逐个跳转；[1-3]/[1,2] 整体链接到第一个编号；
        编号 0（数值区间如 [0,1]，非文献引用）原样输出"""
        for sub in re.split(r'(?<=\])(?=\[)', piece):   # [5][6] → [5]、[6]
            nums = re.findall(r'\d+', sub)
            if not nums or nums[0] == '0':
                self._add_run(par, sub)                 # 数值区间：不上标不链接
                continue
            anchor = f'_Ref{nums[0]}'
            # HYPERLINK 域结构：begin / instr / separate / 显示文本 / end
            self._add_fld(par, 'begin', sup=True)
            self._add_fld(par, 'instr', f' HYPERLINK \\l "{anchor}" ')
            self._add_fld(par, 'separate')
            self._add_run(par, sub, sup=True)           # 显示文本（上标）
            self._add_fld(par, 'end')

    def _add_bookmark(self, par, name):
        """在段落首尾包裹书签（用于正文引用 Ctrl+点击跳转）。
        bookmarkStart 必须位于 pPr 之后（w:p 首个子元素必须是 pPr）"""
        bm_id = str(self._bm_next)
        self._bm_next += 1
        p = par._p
        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), bm_id)
        start.set(qn('w:name'), name)
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            pPr.addnext(start)
        else:
            p.insert(0, start)
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), bm_id)
        p.append(end)

    def _write_rich(self, par, text, sup_refs=True):
        """**加粗**；$...$ 行内公式转 Word 公式；[n] 引用上标"""
        parts = re.split(r'(\$[^$\n]+\$)', text)
        for part in parts:
            if len(part) >= 2 and part.startswith('$') and part.endswith('$'):
                latex = part[1:-1].strip()
                root = _pandoc_parse(latex, block=False)
                children = omml_inline_children(root) if root is not None else []
                if children:
                    for ch in children:
                        par._p.append(ch)
                else:
                    self._add_run(par, part)
                continue
            for sub in re.split(r'(\*\*.+?\*\*)', part):
                if sub.startswith('**') and sub.endswith('**'):
                    self._add_run(par, sub[2:-2], bold=True)
                elif sub:
                    self._emit_refs(par, sub, sup_refs)

    def formula_block(self, latex):
        """居中块公式段落（Word 原生公式）"""
        par = self._new_par()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        root = _pandoc_parse(latex, block=True)
        omml = omml_block_element(root) if root is not None else None
        if omml is not None:
            par._p.append(omml)
        else:
            self._add_run(par, latex)   # pandoc 不可用时兜底显示原样

    def _fill_cell(self, cell, text, bold=False):
        """单元格填充：$...$ 行内公式转 Word 公式、**加粗** 支持；bold=表头整体加粗"""
        p = cell.paragraphs[0]
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        self._write_rich(p, text or '')
        if bold:
            for run in p.runs:
                run.bold = True

    def table(self, rows):
        if not rows:
            return
        self.tab_no += 1
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        tbl = self.doc.add_table(rows=n_rows, cols=n_cols)
        for r, row in enumerate(rows):
            for c in range(n_cols):
                cell = tbl.cell(r, c)
                self._fill_cell(cell, row[c] if c < len(row) else '', bold=(r == 0))
        apply_three_line_borders(tbl)
        self._place(tbl._tbl)
        return tbl

    def image(self, path, caption=None, width_cm=14.5):
        if not os.path.exists(path):
            print(f"[警告] 图片不存在: {path}", file=sys.stderr)
            return
        par = self._new_par()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.add_run().add_picture(path, width=Cm(width_cm))
        if caption:
            self.fig_no += 1
            cap = self._new_par('caption')
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._add_run(cap, f"图{self.fig_no} {caption}")

    # --- 文本文件流式解析 ---
    def _try_block_math(self, lines, i):
        """若 lines[i] 是 $$...$$ 块公式（单行或多行），写入居中公式段落，返回 (新i, True)"""
        s = lines[i].strip()
        if not s.startswith('$$'):
            return i, False
        if s.endswith('$$') and len(s) > 2:
            latex = s[2:-2].strip()
            i += 1
        else:
            buf = [s[2:]]
            i += 1
            n = len(lines)
            while i < n:
                l = lines[i].strip()
                i += 1
                if l.endswith('$$'):
                    buf.append(l[:-2])
                    break
                buf.append(l)
            latex = '\n'.join(buf).strip()
        self.formula_block(latex)
        return i, True

    def write_text_lines(self, lines, path, demote=0, skip_title=None):
        i, n = 0, len(lines)
        skipped = not skip_title
        while i < n:
            line = lines[i].rstrip()
            if not line.strip():
                i += 1
                continue
            if line.strip().startswith('//'):   # 注释行（如分组标记）不打印
                i += 1
                continue
            # 与模板节标题重复的首行 → 跳过
            if not skipped and norm(line) == norm(skip_title):
                skipped = True
                i += 1
                continue
            skipped = True
            i2, handled = self._try_block_math(lines, i)
            if handled:
                i = i2
                continue
            m = MD_IMAGE.match(line.strip())
            if m:
                alt, img = m.group(1), m.group(2)
                img_path = img if os.path.isabs(img) else os.path.join(os.path.dirname(path), img)
                self.image(img_path, caption=alt or None)
                i += 1
                continue
            if MD_TABLE_ROW.match(line):
                rows, i2 = parse_table_rows(lines, i)
                self.table(rows)
                i = i2
                continue
            h = classify_heading(line)
            if h:
                level, text = h
                self.heading(min(level + demote, 3), text, toc_skip=self.toc_skip_headings)
                i += 1
                continue
            if line.strip() in ('---', '***'):
                i += 1
                continue
            s = line.strip()
            # 表注/图注行（表X-X / Table X-X / 图X-X / Fig.X-X）居中显示
            if (re.match(r'^表\s*\d+[-–]\d+', s) or re.match(r'^Table\s*\d+[-–]\d+', s, re.I)
                    or re.match(r'^图\s*\d+[-–]\d+', s) or re.match(r'^Fig(?:ure)?\.?\s*\d+[-–]\d+', s, re.I)):
                self.caption(s)
            else:
                # 行首 [n] 的整行视为参考文献条目（[1]作者.题名[J]...），编号不上标，
                # 并给条目段落加书签 _Ref{n}，供正文引用 Ctrl+点击跳转
                m = re.match(r'^\[(\d+)\]', s)
                if m:
                    par = self.reference(s)             # 参考文献条目：模板格式(12磅/1.25倍行距/悬挂缩进)
                    self._add_bookmark(par, f'_Ref{m.group(1)}')
                else:
                    par = self.body(s)
            i += 1

    def write_with_mounted_tables(self, lines, path, demote=0, skip_title=None):
        """带 csv 表挂载的文本解析：写完「表X-X 表题」表注后，紧跟插入挂载的 csv 表格

        self.table_mounts: {表号('表1-1'): csv路径}
        匹配规则：行内出现 表号 且行首非"Table"（英文表注行跳过重复挂载）
        """
        save_mounts = getattr(self, 'table_mounts', None)
        if save_mounts is None:
            self.write_text_lines(lines, path, demote=demote, skip_title=skip_title)
            return
        mounts = save_mounts          # 直接引用调用方字典，命中即 pop，跨章共享
        self._pending_mounts = mounts
        self._mount_dir = os.path.dirname(path)

        i, n = 0, len(lines)
        skipped = not skip_title
        while i < n:
            line = lines[i].rstrip()
            if not line.strip():
                i += 1
                continue
            if line.strip().startswith('//'):   # 注释行（如分组标记）不打印
                i += 1
                continue
            if not skipped and norm(line) == norm(skip_title):
                skipped = True
                i += 1
                continue
            skipped = True
            i2, handled = self._try_block_math(lines, i)
            if handled:
                i = i2
                continue
            m = MD_IMAGE.match(line.strip())
            if m:
                alt, img = m.group(1), m.group(2)
                img_path = img if os.path.isabs(img) else os.path.join(self._mount_dir, img)
                self.image(img_path, caption=alt or None)
                i += 1
                continue
            if MD_TABLE_ROW.match(line):
                rows, i2 = parse_table_rows(lines, i)
                self.table(rows)
                i = i2
                continue
            h = classify_heading(line)
            if h:
                level, text = h
                self.heading(min(level + demote, 3), text, toc_skip=self.toc_skip_headings)
                i += 1
                continue
            if line.strip() in ('---', '***'):
                i += 1
                continue
            s = line.strip()
            # 表注行检测 + csv 挂载
            # 中文行「表X-X 表题」直接挂载；英文行「Table X-X」仅在对应中文表号
            # 还未被挂载时兜底挂载（覆盖文本只有英文表注的情况）
            mounted = False
            m_cn = re.match(r'^(表\s*\d+[-–]\d+)', s)
            if m_cn:
                key = re.sub(r'\s+', '', m_cn.group(1))
                if key in self._pending_mounts:
                    csv_path, en_caption = self._pending_mounts.pop(key)
                    self.caption(s)                     # 第一行：中文表注（居中）
                    # 第二行：txt 自带英文表注，或 --tables 的「|英文题注」兜底
                    nxt = lines[i + 1].strip() if i + 1 < n else ''
                    if re.match(r'^Table\s*\d+[-–]\d+', nxt, re.I):
                        self.caption(nxt)               # 用 txt 自带的英文表注
                        i += 1
                    elif en_caption:
                        self.caption(f'Table {key[1:]} {en_caption}')
                    self.mount_csv(csv_path)            # 接下来才是表格
                    mounted = True
            elif s.lower().startswith('table'):
                m_en = re.match(r'^Table\s*(\d+)[-–](\d+)', s)
                if m_en:
                    key = f'表{m_en.group(1)}-{m_en.group(2)}'
                    if key in self._pending_mounts:
                        csv_path, en_caption = self._pending_mounts.pop(key)
                        # txt 只有英文表注时，从 csv 文件名兜底生成中文表注，
                        # 保证 中文表注/英文表注/表格 顺序
                        cn = self._cn_caption_from_csv(csv_path, key)
                        if cn:
                            self.caption(cn)
                        self.caption(s)                 # 英文表注行本身（文本已是英文）
                        self.mount_csv(csv_path)
                        mounted = True
            if not mounted:
                if (re.match(r'^表\s*\d+[-–]\d+', s) or re.match(r'^Table\s*\d+[-–]\d+', s, re.I)
                        or re.match(r'^图\s*\d+[-–]\d+', s) or re.match(r'^Fig(?:ure)?\.?\s*\d+[-–]\d+', s, re.I)):
                    self.caption(s)
                else:
                    # 行首 [n] 的整行视为参考文献条目，编号不上标并加书签供正文跳转
                    m = re.match(r'^\[(\d+)\]', s)
                    if m:
                        par = self.reference(s)         # 参考文献条目：模板格式(12磅/1.25倍行距/悬挂缩进)
                        self._add_bookmark(par, f'_Ref{m.group(1)}')
                    else:
                        par = self.body(s)
            i += 1

    def _cn_caption_from_csv(self, csv_path, key):
        """从 csv 文件名兜底生成中文表注：去掉表号前缀与 (中) 后缀"""
        base = os.path.splitext(os.path.basename(csv_path))[0]
        base = re.sub(r'^表\s*\d+[-–]\d+[_\s]*', '', base)
        base = re.sub(r'[（(]\s*中\s*[）)]\s*$', '', base)
        base = base.strip('_ ')
        return f'{key} {base}' if base else ''

    def mount_csv(self, csv_path):
        """读取 csv 并写入表格（不占 tab_no 计数，表号由表注自带）"""
        rows = []
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            for row in csv.reader(f):
                if any(c.strip() for c in row):
                    rows.append([c.strip() for c in row])
        if not rows:
            print(f"[警告] csv 为空: {csv_path}", file=sys.stderr)
            return
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        tbl = self.doc.add_table(rows=n_rows, cols=n_cols)
        for r, row in enumerate(rows):
            for c in range(n_cols):
                cell = tbl.cell(r, c)
                self._fill_cell(cell, row[c] if c < len(row) else '', bold=(r == 0))
        apply_three_line_borders(tbl)
        self._place(tbl._tbl)
        self.tab_no += 1
        return tbl

# ------------------------- 分节解析 -------------------------

def build_section_map(doc):
    """返回按文档顺序的分节列表，每节含: elements/sectPr/break_par/header_text"""
    body = doc.element.body
    sections = []
    current = {'elements': []}
    for child in body.iterchildren():
        if child.tag in (qn('w:p'), qn('w:tbl')):
            current['elements'].append(child)
            if child.tag == qn('w:p'):
                sectPr = child.find(qn('w:pPr') + '/' + qn('w:sectPr'))
                sectPr = child.find('.//' + qn('w:sectPr'))
                if sectPr is not None:
                    current['sectPr'] = sectPr
                    current['break_par'] = child
                    sections.append(current)
                    current = {'elements': []}
        elif child.tag == qn('w:sectPr'):
            current['sectPr'] = child
            current['break_par'] = None  # 文档末尾分节（无段落载体）
            sections.append(current)
            current = {'elements': []}
    # 附加各节页眉文本
    for s in sections:
        s['header_text'] = ''
        for ref in s['sectPr'].findall(qn('w:headerReference')):
            if ref.get(qn('w:type')) == 'default':
                rid = ref.get(qn('r:id'))
                part = doc.part.rels[rid].target_part
                s['header_text'] = par_text_of(part.element).strip()
                s['header_part'] = part
    return sections

SECTION_KEYS = {
    'abstract':    lambda t: '摘' in t and '目录' not in t,
    'abstract_en': lambda t: 'abstract' in t.lower(),
    'toc':         lambda t: '目' in t,
    'conclusion':  lambda t: '结' in t,
    'references':  lambda t: '参考文献' in t,
    'ack':         lambda t: '致' in t or '谢' in t,
}

def find_section(sections, key):
    """按页眉文本定位分节；ack 取最后一节兜底"""
    pred = SECTION_KEYS.get(key)
    if pred:
        for s in sections:
            if pred(s['header_text']):
                return s
    if key == 'ack' and sections:
        return sections[-1]
    return None

def replace_section_content(doc, section, path, table_mounts=None):
    """用文件内容替换分节正文（保留节标题段落与分节符段落）"""
    title_elm = None
    for elm in section['elements']:
        if elm is section.get('break_par'):
            break
        if elm.tag == qn('w:p') and par_text_of(elm).strip():
            title_elm = elm
            break
    title = par_text_of(title_elm).strip() if title_elm is not None else ''

    # 删除标题之后、分节符段落之前的全部内容
    started = False
    to_remove = []
    for elm in section['elements']:
        if elm is title_elm:
            started = True
            continue
        if started:
            if elm is section.get('break_par'):
                break
            to_remove.append(elm)
    for elm in to_remove:
        elm.getparent().remove(elm)

    anchor = section['break_par'] if section.get('break_par') is not None else section['sectPr']
    writer = CursorWriter(doc, mode='before', anchor=anchor)
    if table_mounts:
        writer.table_mounts = table_mounts
        writer.write_with_mounted_tables(read_lines(path), path, skip_title=title)
    else:
        writer.write_text_lines(read_lines(path), path, skip_title=title)
    return writer

# ------------------------- 章节分节插入 -------------------------

def clone_chapter_sectPr(src_sectPr):
    sp = copy.deepcopy(src_sectPr)
    for ref in sp.findall(qn('w:headerReference')):
        sp.remove(ref)
    tp = sp.find(qn('w:titlePg'))
    if tp is not None:
        sp.remove(tp)
    return sp

def make_break_paragraph(sectPr):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pPr.append(sectPr)
    p.append(pPr)
    return p

def insert_block_after(doc, cursor, files, header_title, demote_map=None, src_sectPr=None,
                       table_mounts=None, toc_skip=False):
    """在 cursor 元素之后插入一组文件内容，最后补分节符段落；返回 (新cursor, sectPr元素, 剩余挂载)

    toc_skip=True（附录）：先写「header_title」一级标题（保留在目录中），
    其后所有标题加 outlineLvl=9 不被目录收录
    """
    demote_map = demote_map or {}
    block = CursorWriter(doc, mode='after', anchor=cursor)
    block.toc_skip_headings = toc_skip
    if toc_skip:
        block.heading(1, header_title)      # 附  录（H1，进目录）
    # table_mounts 直接复用并消费（调用方传入全局剩余，命中即弹出）
    pending_mounts = table_mounts if table_mounts is not None else {}
    for path in files:
        ext = os.path.splitext(path)[1].lower()
        demote = demote_map.get(path, 0)
        if ext in TEXT_EXTS:
            if pending_mounts:
                block.table_mounts = pending_mounts
                block.write_with_mounted_tables(read_lines(path), path, demote=demote)
            else:
                block.write_text_lines(read_lines(path), path, demote=demote)
        elif ext in TABLE_EXTS:
            rows = []
            with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                for row in csv.reader(f):
                    rows.append([c.strip() for c in row])
            block.table(rows)
        elif ext in IMG_EXTS:
            cap = os.path.splitext(os.path.basename(path))[0]
            block.image(path, caption=cap)
        else:
            print(f"[警告] 不支持的文件类型，跳过: {path}", file=sys.stderr)
    # 注：剩余挂载由调用方在全部章节完成后统一报告（跨章表注属正常）
    sectPr = clone_chapter_sectPr(src_sectPr)
    brk = make_break_paragraph(sectPr)
    block._place(brk)
    return block.cursor, sectPr, pending_mounts

def set_section_header(doc, sectPr_elm, title, template_header_elm):
    """为分节创建新页眉：克隆模板页眉（含横线样式），文本替换为 title"""
    sec = None
    for s in doc.sections:
        if s._sectPr is sectPr_elm:
            sec = s
            break
    if sec is None:
        print(f"[警告] 未找到分节对象: {title}", file=sys.stderr)
        return
    hdr = sec.header
    hdr.is_linked_to_previous = False   # 创建新页眉部件
    elm = hdr._element
    for ch in list(elm):
        elm.remove(ch)
    for ch in template_header_elm:
        elm.append(copy.deepcopy(ch))
    # 文本替换：保留段落属性（横线），运行只留一个、写入新标题
    sample_rpr = None
    for p in elm.findall(qn('w:p')):
        runs = p.findall(qn('w:r'))
        if runs and sample_rpr is None:
            rpr = runs[0].find(qn('w:rPr'))
            if rpr is not None:
                sample_rpr = copy.deepcopy(rpr)
        for r in runs:
            p.remove(r)
    first_p = elm.find(qn('w:p'))
    if first_p is not None:
        new_r = OxmlElement('w:r')
        if sample_rpr is not None:
            new_r.append(sample_rpr)
        t = OxmlElement('w:t')
        t.text = title
        t.set(qn('xml:space'), 'preserve')
        new_r.append(t)
        first_p.append(new_r)

def enable_update_fields(doc):
    settings = doc.settings.element
    if settings.find(qn('w:updateFields')) is None:
        el = OxmlElement('w:updateFields')
        el.set(qn('w:val'), 'true')
        settings.append(el)

def split_demote(spec):
    """'文件.txt:1' → (文件.txt, 1)"""
    path, _, suffix = spec.rpartition(':')
    if path and suffix.isdigit() and os.path.exists(path):
        return path, int(suffix)
    return spec, 0

# ------------------------- 主流程 -------------------------

def main():
    ap = argparse.ArgumentParser(description='按论文模板分节结构整合内容（覆盖占位+分章页眉）')
    ap.add_argument('--template', required=True, help='Word 模板路径')
    ap.add_argument('--output', required=True, help='输出 docx 路径')
    ap.add_argument('--abstract', default=None, help='中文摘要 txt（替换摘要节占位）')
    ap.add_argument('--abstract-en', default=None, help='英文摘要 txt（替换 Abstract 节占位）')
    ap.add_argument('--conclusion', default=None, help='结论 txt（替换结论节占位）')
    ap.add_argument('--references', default=None, help='参考文献 txt（替换参考文献节占位）')
    ap.add_argument('--acknowledgement', default=None, help='致谢 txt（替换致谢节占位）')
    ap.add_argument('--appendix', default=None, help='附录文件，多个用逗号分隔；:N 后缀=标题降N级')
    ap.add_argument('--tables', default=None, action='append',
                    help='csv表格挂载，可重复：--tables "表1-1=表1-1_xxx.csv|英文题注"；'
                         '写到文本中「表X-X 表题」表注行后自动插入；'
                         '中文表注(居中)+英文表注(居中)+表格 三行结构')
    ap.add_argument('chapters', nargs='*', help='正文章节 txt/md（每个文件=一章，按顺序）')
    ap.add_argument('--img-width', type=float, default=14.5, help='图片宽度(cm)，默认14.5')
    args = ap.parse_args()

    # 解析表格挂载 {表号: (csv路径, 英文题注)}
    table_mounts = {}
    if args.tables:
        for spec in args.tables:
            for item in spec.split(','):
                item = item.strip()
                if not item:
                    continue
                key, _, spec2 = item.partition('=')
                key = re.sub(r'\s+', '', key)
                csv_path, _, en_caption = spec2.strip().partition('|')
                csv_path = csv_path.strip()
                if not csv_path or not os.path.exists(csv_path):
                    print(f"[警告] 挂载表格文件不存在: {item}", file=sys.stderr)
                    continue
                table_mounts[key] = (csv_path, en_caption.strip())
        if table_mounts:
            print(f"[表格挂载] {len(table_mounts)} 个: {', '.join(table_mounts)}")

    if not os.path.exists(args.template):
        sys.exit(f"模板不存在: {args.template}")
    shutil.copyfile(args.template, args.output)
    doc = Document(args.output)

    sections = build_section_map(doc)

    # 1) 覆盖既有节占位文字
    for key, path in [('abstract', args.abstract), ('abstract_en', args.abstract_en),
                      ('conclusion', args.conclusion), ('references', args.references),
                      ('ack', args.acknowledgement)]:
        if not path:
            continue
        sec = find_section(sections, key)
        if sec is None:
            print(f"[警告] 模板中未找到对应分节，跳过: {key}", file=sys.stderr)
            continue
        w = replace_section_content(doc, sec, path)
        print(f"[覆盖] {key} 节 ← {os.path.basename(path)}"
              f"（表格{w.tab_no}个 图片{w.fig_no}张）")

    # 页眉克隆母版：用"摘要"节页眉（横线样式与模板完全一致）
    abstract_sec = find_section(sections, 'abstract')
    header_template = abstract_sec['header_part'].element if abstract_sec and 'header_part' in abstract_sec else None
    if header_template is None and sections:
        for s in sections:
            if 'header_part' in s:
                header_template = s['header_part'].element
                break

    new_headers = []  # (sectPr元素, 页眉标题)

    # 2) 正文章节：每章一个新分节，插在目录节之后
    chapter_groups = []   # 按章分组（文本文件开新章，非文本文件并入当前章）
    for spec in args.chapters:
        path, demote = split_demote(spec)
        ext = os.path.splitext(path)[1].lower()
        if not os.path.exists(path):
            print(f"[警告] 文件不存在，跳过: {path}", file=sys.stderr)
            continue
        if ext in TEXT_EXTS:
            chapter_groups.append({'files': [path], 'demotes': {path: demote}})
        elif chapter_groups:
            chapter_groups[-1]['files'].append(path)
            chapter_groups[-1]['demotes'][path] = demote
        else:
            print(f"[警告] 章节前出现非文本文件，跳过: {path}", file=sys.stderr)

    if chapter_groups:
        toc_sec = find_section(sections, 'toc')
        if toc_sec is None:
            sys.exit("模板中未找到目录分节，无法插入章节")
        cursor = toc_sec['break_par']
        # 分节母版：优先用正文节（结论/摘要）的 sectPr —— 其 docGrid(linePitch=359)
        # 与模板正文一致；目录节 docGrid(linePitch=388) 是目录专用网格，直接克隆
        # 会导致正文字行距/标题段前段后（beforeLines/afterLines 按网格换算）与模板不符
        src_sectPr = None
        for key in ('conclusion', 'abstract'):
            s = find_section(sections, key)
            if s is not None:
                src_sectPr = s['sectPr']
                break
        if src_sectPr is None:
            src_sectPr = toc_sec['sectPr']
        for g in chapter_groups:
            # 首个文本文件的第一个标题 = 章标题（页眉文本）
            title = ''
            for p in g['files']:
                if os.path.splitext(p)[1].lower() in TEXT_EXTS:
                    for line in read_lines(p):
                        if not line.strip():
                            continue
                        h = classify_heading(line)
                        if h and h[0] == 1:
                            title = h[1]
                        break
                    break
            if not title:
                title = os.path.splitext(os.path.basename(g['files'][0]))[0]
            cursor, sp, table_mounts = insert_block_after(doc, cursor, g['files'], title,
                                            demote_map=g['demotes'], src_sectPr=src_sectPr,
                                            table_mounts=table_mounts)
            new_headers.append((sp, title))
            print(f"[章节] {title}（{len(g['files'])}个文件）")

    # 3) 附录：参考文献节之后新建分节
    if args.appendix:
        specs = [s.strip() for s in args.appendix.split(',') if s.strip()]
        files, demotes = [], {}
        for spec in specs:
            p, d = split_demote(spec)
            if os.path.exists(p):
                files.append(p)
                if d:
                    demotes[p] = d
            else:
                print(f"[警告] 附录文件不存在，跳过: {p}", file=sys.stderr)
        if files:
            ref_sec = find_section(sections, 'references')
            if ref_sec is None:
                ref_sec = find_section(sections, 'toc')
            if ref_sec is not None:
                cursor, sp, _ = insert_block_after(doc, ref_sec['break_par'], files,
                                                '附  录', demote_map=demotes,
                                                src_sectPr=ref_sec['sectPr'], toc_skip=True)
                new_headers.append((sp, '附  录'))
                print(f"[附录] {len(files)}个文件")

    # 全部内容写入后：报告未挂载的表格
    if table_mounts:
        for key in sorted(table_mounts):
            print(f"[警告] 表注「{key}」未在文本中找到，未挂载: {table_mounts[key][0]}", file=sys.stderr)

    # 4) 为新分节设置页眉（克隆模板页眉样式）
    if header_template is not None:
        for sp, title in new_headers:
            set_section_header(doc, sp, title, header_template)
        # 修复模板缺陷：致谢节若无自有页眉（原继承上一节），补"致  谢"页眉
        ack_sec_elm = find_section(sections, 'ack')
        if ack_sec_elm is not None and not ack_sec_elm['sectPr'].findall(qn('w:headerReference')):
            ack_title = ''
            for elm in ack_sec_elm['elements']:
                if elm.tag == qn('w:p'):
                    t = par_text_of(elm).strip()
                    if t:
                        ack_title = t
                        break
            if ack_title:
                set_section_header(doc, ack_sec_elm['sectPr'], ack_title, header_template)
                print(f"[页眉修复] 致谢节 ← {ack_title}")
    else:
        print("[警告] 模板无可克隆页眉，新分节将沿用上一节页眉", file=sys.stderr)

    # 5) 打开时自动更新目录域
    enable_update_fields(doc)

    doc.save(args.output)
    print(f"[完成] 已生成: {args.output}")
    print(f"  - 新建分节（含独立页眉）: {len(new_headers)} 个")
    print(f"  - 提示: Word 打开后允许\"更新域\"，目录与页码将自动刷新")

if __name__ == '__main__':
    main()
