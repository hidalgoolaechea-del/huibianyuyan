# -*- coding: utf-8 -*-
"""
汇编语言实验报告生成器（配置驱动）。

用法:
    python generate_report.py <config.json> [-o <output.docx>]

设计约定（按用户需求硬性固定）:
  * 姓名、学号、班级默认留空，只保留带下划线的填空位，由使用者自行填写；
  * 插入截图时不添加任何图注（图 1 / 图 2 之类的说明文字一律不生成）；
  * 版式沿用「实验一/实验二」模板：标题栏 + 信息栏 + 单列边框表格（实验目的 /
    实验内容和原理 / 实验设备与实验环境）+ 实验心得表格。

配置格式见 references/config_schema.md，示例见 assets/report.config.example.json。
"""
import argparse
import json
import os
import sys

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_LINE_SPACING,
                                WD_TAB_ALIGNMENT, WD_TAB_LEADER)
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    sys.exit('缺少 python-docx，请先安装：pip install python-docx')

TEXT_W = Cm(14.6)
FILL = '\u3000' * 6          # 空白填空位（全角空格 + 下划线）


class Builder:
    def __init__(self, cfg, base_dir):
        self.cfg = cfg
        self.base = base_dir
        self.doc = Document()
        self._setup_styles()

    # ---------- 基础样式 ----------
    def _setup_styles(self):
        sec = self.doc.sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.left_margin = Cm(3.17)
        sec.right_margin = Cm(3.17)
        sec.top_margin = Cm(2.54)
        sec.bottom_margin = Cm(2.54)

        normal = self.doc.styles['Normal']
        normal.font.name = '宋体'
        normal.font.size = Pt(10.5)
        normal.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)

    @staticmethod
    def run(run, size=10.5, bold=False, font='宋体', underline=False):
        run.font.name = font
        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline
        return run

    def para(self, container, text='', size=10.5, bold=False,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Pt(21),
             before=0, after=1, line=18, keep_next=False):
        p = container.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.first_line_indent = first_indent
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        if line:
            pf.line_spacing = Pt(line)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.keep_with_next = keep_next
        if text:
            self.run(p.add_run(text), size=size, bold=bold)
        return p

    def code_line(self, container, text):
        p = container.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(30)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(11.5)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        self.run(p.add_run(text), size=9, font='Consolas')
        return p

    def data_line(self, container, ins, res):
        p = container.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(30)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(11.5)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        self.run(p.add_run(ins.ljust(14) + '  '), size=9, font='Consolas')
        self.run(p.add_run(res), size=9)
        return p

    # ---------- 图片（无图注） ----------
    def picture(self, container, filename, width_cm):
        img_dir = self.cfg.get('images_dir', '')
        if not os.path.isabs(img_dir):
            img_dir = os.path.join(self.base, img_dir)
        path = os.path.join(img_dir, filename)
        if not os.path.isfile(path):
            raise FileNotFoundError('截图不存在: ' + path)
        p = container.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(3)
        pf.space_after = Pt(4)
        pf.line_spacing = 1.0
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.add_run().add_picture(path, width=Cm(width_cm))
        # 不生成图注：这是本 skill 的固定约定
        return p

    # ---------- 表格 ----------
    @staticmethod
    def _borders(table, sz=4):
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement('w:' + edge)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), '000000')
            borders.append(el)
        table._tbl.tblPr.append(borders)

    def make_table(self, rows=1):
        t = self.doc.add_table(rows=rows, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        self._borders(t)
        for row in t.rows:
            row.cells[0].width = TEXT_W
        return t

    @staticmethod
    def cell_heading(cell, text):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(2)
        pf.space_after = Pt(3)
        pf.line_spacing = Pt(22)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        Builder.run(p.add_run(text), size=14)

    # ---------- 信息栏 ----------
    def info_line(self, label, value, tab_pt):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(30)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.tab_stops.add_tab_stop(tab_pt, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
        self.run(p.add_run(label), size=12)
        self.run(p.add_run(value if value else FILL), size=12, underline=True)
        p.add_run('\t')
        return p

    def pair_line(self, label1, value1, label2, value2):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.line_spacing = Pt(30)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.tab_stops.add_tab_stop(Cm(7.55), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
        pf.tab_stops.add_tab_stop(Cm(14.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
        self.run(p.add_run(label1), size=12)
        self.run(p.add_run(value1 if value1 else FILL), size=12, underline=True)
        p.add_run('\t')
        self.run(p.add_run(label2), size=12)
        p.add_run('\t')
        self.run(p.add_run(value2 if value2 else FILL), size=12, underline=True)
        return p

    # ---------- 组装 ----------
    def build(self):
        c = self.cfg
        # 标题
        t = self.doc.add_paragraph()
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.paragraph_format.space_before = Pt(0)
        t.paragraph_format.space_after = Pt(10)
        t.paragraph_format.line_spacing = Pt(30)
        t.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        self.run(t.add_run(c.get('school_header', '计电学院《汇编语言》实验报告')),
                 size=16, bold=True)

        # 信息栏（姓名/学号/班级留空）
        self.info_line('实验项目    ', c.get('project', ''), Cm(14.6))
        self.info_line('实验日期    ', c.get('date', ''), Cm(6.2))
        self.pair_line('班    级    ', c.get('class_name', ''),
                       '  姓    名', c.get('name', ''))
        self.pair_line('学    号    ', c.get('student_id', ''),
                       '  综合成绩', c.get('score', ''))

        # 一、实验内容
        h = self.doc.add_paragraph()
        h.paragraph_format.first_line_indent = Pt(0)
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(4)
        self.run(h.add_run('一、实验内容'), size=14, bold=True)

        table = self.make_table(3)
        cell = table.cell(0, 0)
        self.cell_heading(cell, '1、实验目的')
        for g in c.get('goals', []):
            self.para(cell, g, line=18, after=0)

        cell = table.cell(1, 0)
        self.cell_heading(cell, '2、实验内容和原理')
        for i, sec in enumerate(c.get('sections', [])):
            self.para(cell, sec.get('title', ''), bold=True,
                      before=2 if i == 0 else 6, after=1, keep_next=True)
            for blk in sec.get('blocks', []):
                self.render_block(cell, blk)

        cell = table.cell(2, 0)
        self.cell_heading(cell, '3、实验设备与实验环境')
        for e in c.get('environment', self.default_env(c)):
            self.para(cell, e, line=18, after=0)

        # 二、实验心得
        h2 = self.doc.add_paragraph()
        h2.paragraph_format.first_line_indent = Pt(0)
        h2.paragraph_format.space_before = Pt(10)
        h2.paragraph_format.space_after = Pt(4)
        self.run(h2.add_run('二、实验心得'), size=14, bold=True)
        t2 = self.make_table(1)
        cell = t2.cell(0, 0)
        for tx in c.get('thoughts', []):
            self.para(cell, tx, line=19, after=1)
        return self.doc

    @staticmethod
    def default_env(c):
        return [
            '（1）x86-64 个人计算机，Windows 11 操作系统；',
            '（2）DOSBox 0.74 模拟器与 MASM 5.0 工具包（MASM.EXE、LINK.EXE、DEBUG.EXE）；',
            '（3）教材：《汇编语言（第 2 版）》（王爽），%s。' % c.get('lab_reference', ''),
        ]

    def render_block(self, cell, blk):
        kind = blk.get('type', 'text')
        if kind == 'text':
            self.para(cell, blk.get('value', ''),
                      bold=blk.get('bold', False),
                      before=blk.get('before', 0),
                      after=blk.get('after', 1))
        elif kind == 'code':
            for line in blk.get('lines', []):
                self.code_line(cell, line)
        elif kind == 'data':
            for row in blk.get('rows', []):
                self.data_line(cell, row[0], row[1])
        elif kind == 'image':
            self.picture(cell, blk['file'],
                         blk.get('width_cm', self.cfg.get('image_width_cm', 11.8)))
        else:
            raise ValueError('未知区块类型: ' + kind)


def main():
    ap = argparse.ArgumentParser(description='生成汇编语言实验报告 docx')
    ap.add_argument('config', help='配置 JSON 路径')
    ap.add_argument('-o', '--output', help='覆盖配置中的 output 路径')
    args = ap.parse_args()

    cfg_path = os.path.abspath(args.config)
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    out = args.output or cfg.get('output')
    if not out:
        sys.exit('配置缺少 output 字段，或用 -o 指定输出路径')
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out), exist_ok=True)

    builder = Builder(cfg, os.path.dirname(cfg_path))
    builder.build().save(out)
    print('OK:', out, os.path.getsize(out), 'bytes')


if __name__ == '__main__':
    main()
