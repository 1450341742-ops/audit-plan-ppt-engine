from __future__ import annotations
import io
import zipfile
import traceback
from pathlib import Path
import streamlit as st
from batch_generate import render_one
from ai_summary import _get_cfg, get_last_ai_status

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "output"
ASSETS_DIR = BASE_DIR / "assets"
TEMPLATE_PATH = ASSETS_DIR / "template.pptx"
OUT_DIR.mkdir(parents=True, exist_ok=True)
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="稽查总结会PPT生成器 V8.12", layout="centered")
st.title("稽查总结会PPT生成器 V8.12（内置模板版）")
st.caption("系统已内置稽查总结会PPT模板，只需上传Excel表格即可生成PPT。TOP5页优先调用AI总结；AI未配置或调用失败时自动使用规则聚类兜底。")

repo_template_ok = TEMPLATE_PATH.exists() and TEMPLATE_PATH.stat().st_size > 1024 * 100
if repo_template_ok:
    st.success(f"已检测到内置模板：assets/template.pptx（{TEMPLATE_PATH.stat().st_size/1024/1024:.1f} MB）")
else:
    st.error("未检测到有效内置模板：assets/template.pptx。请先将新版稽查总结会模板放入 assets/template.pptx 后再部署。")

coze_key = _get_cfg("COZE_API_KEY") or _get_cfg("COZE_TOKEN")
coze_bot_id = _get_cfg("COZE_BOT_ID")
coze_base_url = _get_cfg("COZE_BASE_URL") or "https://api.coze.cn"
openai_key = _get_cfg("OPENAI_API_KEY") or _get_cfg("DINGTALK_API_KEY") or _get_cfg("DEAP_API_KEY") or _get_cfg("AI_API_KEY")
openai_model = _get_cfg("OPENAI_MODEL") or _get_cfg("DINGTALK_MODEL") or _get_cfg("DEAP_MODEL")

with st.expander("AI配置状态", expanded=True):
    if coze_key and coze_bot_id:
        st.success("扣子AI配置已检测到：COZE_API_KEY/COZE_TOKEN + COZE_BOT_ID")
        st.caption(f"COZE_BASE_URL：{coze_base_url}")
        st.caption(f"COZE_BOT_ID尾号：{coze_bot_id[-6:] if len(coze_bot_id) >= 6 else coze_bot_id}")
    elif openai_key and openai_model:
        st.success("OpenAI/兼容AI配置已检测到")
        st.caption(f"MODEL：{openai_model}")
    else:
        st.warning("未检测到完整AI配置。TOP5页会使用规则聚类兜底。")
        st.markdown(
            """
            扣子接入至少需要在 Streamlit Secrets 中配置：
            ```toml
            COZE_API_KEY = "你的扣子Secret token"
            COZE_BOT_ID = "你的扣子Bot ID"
            COZE_BASE_URL = "https://api.coze.cn"
            ```
            """
        )
    st.info("生成PPT后，TOP5页只展示高风险问题、风险维度分析和核查应对建议，不再显示生成来源说明。")

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
        status = get_last_ai_status()
        if status.get("ok"):
            st.success(f"AI调用成功：{status.get('source')}｜{status.get('message')}")
        else:
            st.warning(f"AI未成功使用：{status.get('source')}｜{status.get('message')}")
        if outs:
            st.success("生成完成。请下载PPT查看TOP5页内容。")
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
