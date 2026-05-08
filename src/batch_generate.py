from __future__ import annotations
import argparse
from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from parser import parse_excel
import renderer

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OUTPUT = BASE_DIR / "output"
DEFAULT_TEMPLATE = BASE_DIR / "assets" / "template.pptx"
DEFAULT_OUTPUT.mkdir(exist_ok=True, parents=True)

DARK_BLUE = RGBColor(31, 78, 121)
FIXED_COVER_COMPANY = "北京xxxx医药科技有限公司"


def safe_stem(name: str) -> str:
    invalid = '<>:"/\\|?*'
    return "".join("_" if c in invalid else c for c in name).strip().strip(".") or "output"


def _patched_render_cover(slide, context):
    renderer._remove_text_shapes(slide)
    meta = context.get("meta", {})
    project = renderer._clean(meta.get("project_name", "—"))
    center = renderer._clean(meta.get("center_name", "—"))
    center_no = renderer._clean(meta.get("center_no", ""))
    audit_date = renderer._clean(meta.get("audit_date", "—"))

    title = f"{project}-{center}" + (f"（中心编号{center_no}）" if center_no else "")
    line_count = renderer._estimate_lines(title, 35)
    title_font = 28
    if line_count >= 4:
        title_font = 22
    elif line_count == 3:
        title_font = 24
    elif line_count == 2:
        title_font = 26

    renderer._add_textbox(slide, 0.25, 2.88, 12.70, 1.50, title, title_font, True, renderer.WHITE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.25, 4.58, 12.70, 0.44, "中心稽查末次会议", 28, True, renderer.YELLOW, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.35, 5.24, 12.30, 0.30, f"时间：{audit_date}", 14, True, DARK_BLUE, PP_ALIGN.LEFT)
    renderer._add_textbox(slide, 0.35, 5.56, 12.30, 0.30, FIXED_COVER_COMPANY, 14, True, DARK_BLUE, PP_ALIGN.LEFT)


renderer._render_cover = _patched_render_cover


def render_one(excel_path: str | Path, output_dir: str | Path | None = None, template_path: str | Path | None = None) -> Path:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir or DEFAULT_OUTPUT)
    output_dir.mkdir(parents=True, exist_ok=True)
    context = parse_excel(excel_path)
    out = output_dir / f"{safe_stem(excel_path.stem)}-V8.4封面修正版.pptx"
    renderer.render_ppt(context, out, template_path=Path(template_path or DEFAULT_TEMPLATE))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", type=str)
    ap.add_argument("--excel_dir", type=str)
    ap.add_argument("--output_dir", type=str, default=str(DEFAULT_OUTPUT))
    ap.add_argument("--template", type=str, default=str(DEFAULT_TEMPLATE))
    args = ap.parse_args()
    files = []
    if args.excel:
        files.append(Path(args.excel))
    if args.excel_dir:
        d = Path(args.excel_dir)
        for ext in ("*.xlsx", "*.xlsm", "*.xls"):
            files += sorted(d.glob(ext))
    if not files:
        raise SystemExit("请提供 --excel 或 --excel_dir")
    for f in files:
        out = render_one(f, args.output_dir, template_path=args.template)
        print(out)


if __name__ == "__main__":
    main()
