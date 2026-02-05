#!/usr/bin/env python
# Copyright (c) 2020 Kevin McGuinness <kevin.mcguinness@gmail.com>

import click
import sys
import os
from pathlib import Path

from . import convert_pdf2pptx


arg = click.argument
opt = click.option


@click.command()
@opt('-o', '--output', 'output_file', default=None,
     help='location to save the pptx (default: PDF_FILE.pptx or output directory for folder input)')
@opt('-r', '--resolution', default=300, type=int,
     help='resolution in dots per inch (default: 300)')
@opt('-q', '--quiet', is_flag=True, default=False,
     help='disable printing progress bar and other info')
@opt('--from', 'start_page', default=0, type=int,
     help='first page in the pdf to copy to the pptx (only for single file)')
@opt('--count', 'page_count', default=None, type=int,
     help='number of pages in the pdf to copy to the pptx (only for single file)')
@arg('pdf_file', type=click.Path(exists=True, dir_okay=True))
def main(pdf_file, output_file, resolution, start_page, page_count, quiet):
    """
    Convert a PDF slideshow to Powerpoint PPTX.

    Renders each page as a PNG image and creates the resulting Powerpoint 
    slideshow from these images. Useful when you want to use Powerpoint
    to present a set of PDF slides (e.g. slides from Beamer). You can then
    use the presentation capabilities of Powerpoint (notes, ink on slides,
    etc.) with slides created in LaTeX.
    
    If pdf_file is a directory, all PDF files in the directory will be converted.
    """
    try:
        # 检查输入是否为目录
        if os.path.isdir(pdf_file):
            # 处理目录输入
            input_dir = Path(pdf_file)
            pdf_files = sorted(input_dir.glob('*.pdf'))
            
            if not pdf_files:
                print(f"错误：目录 '{pdf_file}' 中未找到 PDF 文件", file=sys.stderr)
                sys.exit(1)
            
            # 检查是否使用了单文件专用参数
            if start_page != 0 or page_count is not None:
                print("警告：--from 和 --count 参数仅适用于单个文件，处理目录时将被忽略", file=sys.stderr)
            
            # 确定输出目录
            if output_file:
                output_dir = Path(output_file)
                output_dir.mkdir(parents=True, exist_ok=True)
            else:
                output_dir = input_dir
            
            # 处理每个 PDF 文件
            total = len(pdf_files)
            if not quiet:
                print(f"找到 {total} 个 PDF 文件")
            
            for idx, pdf_path in enumerate(pdf_files, 1):
                if not quiet:
                    print(f"\n[{idx}/{total}] 正在处理: {pdf_path.name}")
                
                output_path = output_dir / pdf_path.with_suffix('.pptx').name
                convert_pdf2pptx(
                    str(pdf_path), str(output_path), resolution, 0, None, quiet)
            
            if not quiet:
                print(f"\n完成！已转换 {total} 个文件")
        else:
            # 处理单个文件输入（原有逻辑）
            convert_pdf2pptx(
                pdf_file, output_file, resolution, start_page, page_count, quiet)
    except PermissionError as err:
        print(err, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter
