from __future__ import annotations
import io
import zipfile
import traceback
from pathlib import Path
import streamlit as st
from batch_generate import render_one

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "output"
ASSETS_DIR = BASE_DIR / "assets"
TEMPLATE_PATH = ASSETS_DIR / "template.pptx"
OUT_DIR.mkdir(parents=True, exist_ok=True)
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="稽查总结会PPT生成器 V8.4", layout="centered")
st.title("稽查总结会PPT生成器 V8.4（内置模板版）")
st.caption("系统已内置稽查总结会PPT模板，只需上传Excel表格即可生成PPT。共性问题、个性问题、Q&A 三页仅保留模板，不写入内容；已取消建议项页。")

repo_template_ok = TEMPLATE_PATH.exists() and TEMPLATE_PATH.stat().st_size > 1024 * 100
if repo_template_ok:
    st.success(f"已检测到内置模板：assets/template.pptx（{TEMPLATE_PATH.stat().st_size/1024/1024:.1f} MB）")
else:
    st.error("未检测到有效内置模板：assets/template.pptx。请先将新版稽查总结会模板放入 assets/template.pptx 后再部署。")

uploads = st.file_uploader("上传Excel文件", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True)


def zip_bytes(paths: list[Path]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            if p.exists():
                zf.write(p, arcname=p.name)
    return bio.getvalue()


if st.button("开始生成", type="primary"):
    template_ok = TEMPLATE_PATH.exists() and TEMPLATE_PATH.stat().st_size > 1024 * 100
    if not template_ok:
        st.error("系统未检测到有效内置模板，无法生成。请确认仓库中存在 assets/template.pptx。")
    elif not uploads:
        st.warning("请上传 Excel 文件")
    else:
        outs = []
        for file_idx, f in enumerate(uploads):
            in_path = OUT_DIR / f"{file_idx}_{f.name}"
            in_path.write_bytes(f.getvalue())
            try:
                out = render_one(in_path, OUT_DIR, template_path=TEMPLATE_PATH)
                outs.append(Path(out))
            except Exception as e:
                st.error(f"{f.name} 生成失败：{e}")
                st.code(traceback.format_exc(), language="python")
        if outs:
            st.success("生成完成")
            for idx, p in enumerate(outs):
                with open(p, "rb") as fp:
                    st.download_button(
                        f"下载 {p.name}",
                        data=fp.read(),
                        file_name=p.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"download_{idx}_{p.name}",
                    )
            st.download_button(
                "下载全部PPT（ZIP）",
                data=zip_bytes(outs),
                file_name="audit_summary_ppt_results.zip",
                mime="application/zip",
                key=f"zip_all_{len(outs)}",
            )
