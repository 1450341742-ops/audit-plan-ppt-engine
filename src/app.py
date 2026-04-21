from __future__ import annotations
import io
import zipfile
import traceback
from pathlib import Path
import streamlit as st
from batch_generate import render_one

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="稽查总结会PPT生成器 V8.0", layout="centered")
st.title("稽查总结会PPT生成器 V8.0")
st.caption("Streamlit Cloud / Windows / Mac 通用版：使用 python-pptx 生成标准 .pptx，不依赖 pywin32 或本机 PowerPoint。")
st.success("当前版本支持部署到 Streamlit Cloud，生成的PPT可用 Microsoft PowerPoint、WPS、Mac PowerPoint 打开。")

uploads = st.file_uploader("上传Excel文件", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True)


def zip_bytes(paths: list[Path]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            if p.exists():
                zf.write(p, arcname=p.name)
    return bio.getvalue()


if st.button("开始生成", type="primary"):
    if not uploads:
        st.warning("请先上传 Excel 文件")
    else:
        outs = []
        for f in uploads:
            in_path = OUT_DIR / f.name
            in_path.write_bytes(f.getvalue())
            try:
                out = render_one(in_path, OUT_DIR)
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
