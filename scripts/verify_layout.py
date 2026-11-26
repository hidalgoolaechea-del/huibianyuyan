# -*- coding: utf-8 -*-
"""
报告排版检查：用 Word 导出 PDF，统计页数，并生成整页缩略拼图供人工/视觉核验。

用法:
    python verify_layout.py <报告1.docx> [报告2.docx ...] [--out sheet.png] [--dpi 50]

注意:
  * Word COM 无法打开文件名含方括号（如 [bx]）的文档，脚本会自动复制到临时目录并改名；
  * 若上一次检查异常退出，可能有 WINWORD.EXE 残留锁住文件，先结束该进程再重试；
  * 依赖: pywin32, pymupdf, pillow（仅 Windows 可用 Word COM）。
"""
import argparse
import os
import shutil
import sys
import tempfile

try:
    import win32com.client as win32
    import pymupdf
    from PIL import Image
except ImportError as e:
    sys.exit('缺少依赖 (%s)：pip install pywin32 pymupdf pillow' % e)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docs', nargs='+')
    ap.add_argument('--out', default='report_sheet.png')
    ap.add_argument('--dpi', type=int, default=50)
    args = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix='asmrep_')
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    thumbs = []
    try:
        for i, path in enumerate(args.docs):
            path = os.path.abspath(path)
            if not os.path.isfile(path):
                print('跳过（不存在）:', path)
                continue
            # 文件名含 [ ] 等字符时 Word COM 会做 URL 转义导致找不到文件
            open_path = path
            if any(ch in os.path.basename(path) for ch in '[]%#'):
                open_path = os.path.join(tmp, 'doc%d.docx' % i)
                shutil.copy2(path, open_path)

            pdf = os.path.join(tmp, 'doc%d.pdf' % i)
            d = word.Documents.Open(open_path, False, False)
            try:
                d.ExportAsFixedFormat(pdf, 17)
                pages = d.ComputeStatistics(2)
            finally:
                d.Close(False)
            print('%s -> %d 页' % (os.path.basename(path), pages))

            doc = pymupdf.open(pdf)
            for pno in range(len(doc)):
                pix = doc[pno].get_pixmap(dpi=args.dpi)
                thumbs.append(Image.frombytes('RGB', (pix.width, pix.height), pix.samples))
            doc.close()
    finally:
        word.Quit()

    if not thumbs:
        print('没有可检查的页面')
        return
    w = max(t.width for t in thumbs)
    h = max(t.height for t in thumbs)
    cols = min(4, len(thumbs))
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * w, rows * h), 'white')
    for i, t in enumerate(thumbs):
        sheet.paste(t, ((i % cols) * w, (i // cols) * h))
    out = os.path.abspath(args.out)
    sheet.save(out)
    print('拼图:', out, sheet.size, '（共 %d 页）' % len(thumbs))


if __name__ == '__main__':
    main()
