from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import load_config
from .engine import collect_docx, refresh_document, refresh_many


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="把 Word 正文和表格内文本的排版一键统一")
    parser.add_argument("input", help="docx 文件或文件夹")
    parser.add_argument("-o", "--output", help="输出文件（单文件）或输出文件夹（批量）")
    parser.add_argument("-c", "--config", help="config.yaml 路径")
    parser.add_argument("--in-place", action="store_true", help="覆盖原文件（请先自行备份）")
    args = parser.parse_args(argv)

    src = Path(args.input).expanduser().resolve()
    cfg_path = Path(args.config).expanduser() if args.config else None
    cfg = load_config(cfg_path)

    try:
        files = collect_docx(src)
    except ValueError as exc:
        print(exc, file=sys.stderr)
        return 2

    if len(files) == 1 and not src.is_dir():
        if args.in_place:
            out = files[0]
        elif args.output:
            out = Path(args.output).expanduser()
        else:
            out = None
        report = refresh_document(files[0], out, cfg)
    else:
        out_dir = Path(args.output).expanduser() if args.output else (src if src.is_dir() else src.parent)
        if args.in_place:
            # 批量原地：逐个覆盖
            from .engine import RefreshReport

            report = None
            for path in files:
                one = refresh_document(path, path, cfg)
                if report is None:
                    report = one
                    report.files = 1
                else:
                    report.merge(one)
        else:
            report = refresh_many(files, out_dir, cfg)

    print(report.as_text())
    return 1 if report.errors else 0


if __name__ == "__main__":
    sys.exit(main())
