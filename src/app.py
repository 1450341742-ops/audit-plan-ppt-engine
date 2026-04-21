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

st.set_page_config(page_title="稽查总结会PPT生成器 V8.2", layout="centered")
st.title("稽查总结会PPT生成器 V8.2（严格模板上传版）")
st.caption("必须使用你上传的 PPT 模板生成：先上传 template.pptx，再上传 Excel。支持 Streamlit Cloud / Windows / Mac，生成标准 .pptx，可用 PowerPoint、WPS 打开。")

repo_template_ok = TEMPLATE_PATH.exists() and TEMPLATE_PATH.stat().st_size > 1024 * 100
if repo_template_ok:
    st.success(f"已检测到仓库模板：assets/template.pptx（{TEMPLATE_PATH.stat().st_size/1024/1024:.1f} MB）")
else:
    st.warning("当前仓库中的 assets/template.pptx 不存在或不是有效模板。请先上传你的稽查总结会PPT模板。")

template_upload = st.file_uploader("① 上传PPT模板（必须上传你的稽查总结会模板，文件名可任意）", type=["pptx"], accept_multiple_files=False)
if template_upload is not None:
    TEMPLATE_PATH.write_bytes(template_upload.getvalue())
    st.success(f"模板已加载：{template_upload.name}（{TEMPLATE_PATH.stat().st_size/1024/1024:.1f} MB）。本次生成将严格使用该模板。")

uploads = st.file_uploader("② 上传Excel文件", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True)


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
        st.error("请先上传有效的 PPT 模板。当前模板为空或无效，无法严格按模板生成。")
    elif not uploads:
        st.warning("请上传 Excel 文件")
    else:
        outs = []
        for f in uploads:
            in_path = OUT_DIR / f.name
            in_path.write_bytes(f.getvalue())
            try:
                out = render_one(in_path, OUT_DIR, template_path=TEMPLATE_PATH)
                outs.append(Path(out))
            except Exception as e:
                st.error(f"{f.name} 生成失败：{e}")
                st.code(traceback.format_exc(), language="python")
        if outs:
            st.success("生成完成")
            for p in outs:
                with open(p, "rb") as fp:
                    st.download_button(
                        f"下载 {p.name}",
                        data=fp.read(),
                        file_name=p.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=p.name,
                    )
            st.download_button(
                "下载全部PPT（ZIP）",
                data=zip_bytes(outs),
                file_name="稽查总结会PPT生成结果.zip",
                mime="application/zip",
                key="zip_all",
            )
