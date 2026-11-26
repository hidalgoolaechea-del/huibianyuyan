# -*- coding: utf-8 -*-
"""
DOSBox / DEBUG 截图预处理：裁掉底部黑边，按 NEAREST 整数倍放大（保持字符锐利）。

用法:
    python prepare_images.py <源截图目录> <输出目录> [--scale 2]

依赖: Pillow（pip install pillow）
"""
import argparse
import os
import sys

try:
    from PIL import Image
except ImportError:
    sys.exit('缺少 Pillow，请先安装：pip install pillow')


def process(src, dst, scale):
    os.makedirs(dst, exist_ok=True)
    names = sorted(n for n in os.listdir(src) if n.lower().endswith('.png'))
    if not names:
        print('警告：源目录中没有 PNG 文件:', src)
    for name in names:
        im = Image.open(os.path.join(src, name)).convert('RGB')
        w, h = im.size
        px = im.load()
        bottom = h
        # 从下往上找最后一行有内容的像素（每 7 列采样一次，避免噪点）
        for y in range(h - 1, -1, -1):
            row = [px[x, y] for x in range(0, w, 7)]
            if all(sum(c) < 40 for c in row):
                bottom = y
            else:
                break
        im = im.crop((0, 0, w, max(bottom, 60)))
        im = im.resize((im.width * scale, im.height * scale), Image.NEAREST)
        im.save(os.path.join(dst, name))
        print(name, im.size)
    print('done ->', dst)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--scale', type=int, default=2)
    args = ap.parse_args()
    process(args.src, args.dst, args.scale)


if __name__ == '__main__':
    main()
