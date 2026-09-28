#!/usr/bin/env python3
"""单一 Typst 内容源的 PDF / Word 导出入口。"""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]

def find_typst():
    candidates = [os.environ.get('TYPST_BIN'), str(ROOT/'.tools'/'typst.exe'), shutil.which('typst')]
    for value in candidates:
        if value and Path(value).is_file(): return str(Path(value).resolve())
    raise RuntimeError('未找到 Typst。请运行 setup.ps1，或安装 Typst 0.15.1 并设置 TYPST_BIN。')

def run(args):
    p=subprocess.run(args,cwd=ROOT,encoding='utf-8',capture_output=True)
    if p.stderr: print(p.stderr,file=sys.stderr,end='')
    if p.returncode: raise RuntimeError(f'命令执行失败（{p.returncode}）：{args[0]} {args[1]}')
    return p.stdout

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--format',choices=['all','pdf','docx'],default='all')
    ap.add_argument('--version',type=int,default=1,help='输出文件名前的 v 版本号')
    ap.add_argument('--update-fields',action='store_true',help='使用本机 Microsoft Word 更新目录和页码')
    args=ap.parse_args()
    if args.version<1: ap.error('版本号必须大于0')
    typst=find_typst()
    print(run([typst,'--version']).strip())
    cfg=json.loads((ROOT/'layout.json').read_text(encoding='utf-8-sig'))
    fonts=run([typst,'fonts'])
    missing=[cfg[k] for k in ['body_font','heading_font','ack_font','latin_font'] if cfg[k].lower() not in fonts.lower()]
    if missing: raise RuntimeError('缺少规范字体：'+', '.join(missing)+'。请合法安装字体，或明确修改 layout.json 使用其他字体后重新验收。')
    raw=run([typst,'eval','query(<thesis-data>).first().value','--in','main.typ','--root',str(ROOT)])
    data=json.loads(raw)
    from export_docx import validate,build_docx
    validate(data,ROOT)
    target=ROOT/'build';target.mkdir(exist_ok=True)
    (target/'content.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    stem=f'v{args.version}_本科毕业论文模板'
    outputs=[]
    if args.format in ('all','pdf'):
        pdf=target/(stem+'.pdf');run([typst,'compile','main.typ',str(pdf),'--root',str(ROOT)]);outputs.append(pdf)
    if args.format in ('all','docx'):
        docx=target/(stem+'.docx');build_docx(data,cfg,ROOT,docx);outputs.append(docx)
        if args.update_fields:
            if sys.platform!='win32': print('提示：当前非 Windows，请在 Word 中更新整个目录。')
            else:
                try: run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(ROOT/'scripts'/'update-word-fields.ps1'),'-DocumentPath',str(docx)])
                except RuntimeError as ex: print('目录自动更新未完成：'+str(ex)+'。DOCX 已生成，可在 Word 中右键更新整个目录。',file=sys.stderr)
    # Word 会填入本机最后修改者；导出时清除该可识别元数据。
    if args.format in ('all','docx'):
        from docx import Document
        exported=Document(docx)
        exported.core_properties.last_modified_by=''
        exported.core_properties.comments=''
        exported.save(docx)
    for p in outputs:print(f'已生成：{p}（{p.stat().st_size:,} 字节）')
    print('本模板含待填写项。正式提交前请检查学院最新要求、目录与每页图表。')

if __name__=='__main__':
    try: main()
    except (RuntimeError,ValueError,ImportError,KeyError) as ex:
        print('导出失败：'+str(ex),file=sys.stderr);sys.exit(1)
