from __future__ import annotations

import json
import os
import re
import time
from typing import Any

LAST_AI_STATUS = {
    "source": "未运行",
    "ok": False,
    "message": "尚未执行AI总结",
}


def _set_status(source: str, ok: bool, message: str) -> None:
    LAST_AI_STATUS["source"] = source
    LAST_AI_STATUS["ok"] = ok
    LAST_AI_STATUS["message"] = message
    _log(f"{source}｜{message}")


def get_last_ai_status() -> dict:
    return dict(LAST_AI_STATUS)


def _log(message: str) -> None:
    print(f"[AI_SUMMARY] {message}", flush=True)


def _clean(value: Any) -> str:
    text = str(value or "").replace("\r", "\n")
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())


def _get_cfg(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value:
        return str(value).strip()
    try:
        import streamlit as st

        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:
        pass
    return default


def _compact_issues(context: dict, max_items: int = 80) -> str:
    lines = []
    for idx, issue in enumerate(context.get("issues", [])[:max_items], start=1):
        category = _clean(issue.get("category", ""))
        severity = _clean(issue.get("severity", ""))
        summary = _clean(issue.get("summary", ""))
        desc = _clean(issue.get("description", ""))
        basis = _clean(issue.get("basis", ""))
        item = f"{idx}. 分类：{category}\n严重程度：{severity}\n问题摘要：{summary}\n问题描述：{desc}\n依据：{basis}"
        lines.append(item[:1200])
    return "\n\n".join(lines)


def _normalize_rows(data: Any) -> list[dict]:
    """把AI返回的多种字段格式统一为PPT可写入的TOP5结构。"""
    if isinstance(data, dict):
        for key in ("items", "top5", "data", "results", "list"):
            if isinstance(data.get(key), list):
                data = data[key]
                break

    if not isinstance(data, list):
        return []

    rows = []
    for item in data[:5]:
        if not isinstance(item, dict):
            continue
        risk = _clean(
            item.get("高风险问题")
            or item.get("risk")
            or item.get("问题")
            or item.get("title")
            or item.get("name")
        )
        analysis = _clean(
            item.get("风险维度分析")
            or item.get("analysis")
            or item.get("风险分析")
            or item.get("reason")
        )
        advice = _clean(
            item.get("核查应对建议")
            or item.get("advice")
            or item.get("建议")
            or item.get("actions")
        )
        if not risk or not analysis or not advice:
            continue
        rows.append(
            {
                "risk": risk[:180],
                "analysis": analysis[:360],
                "advice": advice[:420],
                "score": 999,
                "source": "AI智能总结",
            }
        )
    return rows


def _strip_code_fence(text: str) -> str:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


def _safe_parse_json(text: str) -> list[dict]:
    text = _strip_code_fence(text)
    if not text:
        _set_status("AI解析", False, "AI返回为空")
        return []

    # 1）优先直接解析完整JSON。
    try:
        rows = _normalize_rows(json.loads(text))
        if rows:
            return rows
    except Exception:
        pass

    # 2）兼容扣子/大模型偶尔在文本前后追加说明的情况，提取JSON数组或JSON对象。
    candidates = []
    arr_start, arr_end = text.find("["), text.rfind("]")
    if arr_start >= 0 and arr_end > arr_start:
        candidates.append(text[arr_start : arr_end + 1])
    obj_start, obj_end = text.find("{"), text.rfind("}")
    if obj_start >= 0 and obj_end > obj_start:
        candidates.append(text[obj_start : obj_end + 1])

    for candidate in candidates:
        try:
            rows = _normalize_rows(json.loads(candidate))
            if rows:
                return rows
        except Exception:
            continue

    _set_status("AI解析", False, f"JSON解析失败或字段不匹配，AI返回前300字：{text[:300]}")
    return []


def _system_prompt() -> str:
    return """
你是临床试验第三方稽查与注册核查准备专家。请严格基于输入的中心稽查发现，提炼TOP5高风险问题。
必须只输出JSON数组，正好5项，不要Markdown，不要解释。每项字段必须为：高风险问题、风险维度分析、核查应对建议。
要求：
1. 高风险问题：概括成核查风险主题，不照搬原文。
2. 风险维度分析：写清数据可靠性、受试者安全、方案依从性、受试者权益、伦理合规、药品/样本链条、疗效评价、注册核查等维度及原因。
3. 核查应对建议：按“立即行动/证据准备/系统改进或演练”写，必须可执行。
4. 5项之间不能重复，每个字段适合PPT表格展示。
""".strip()


def _user_prompt(context: dict) -> str:
    return f"请根据以下中心稽查发现，生成TOP5高风险问题及核查应对建议。\n\n{_compact_issues(context)}"


def _extract_messages_data(raw: Any) -> list[dict]:
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        data = raw.get("data")
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("items", "messages", "list"):
                if isinstance(data.get(key), list):
                    return data[key]
    return []


def _coze_error_message(payload: Any) -> str:
    if not isinstance(payload, dict):
        return str(payload)[:500]
    code = payload.get("code")
    msg = payload.get("msg") or payload.get("message") or payload.get("error_message") or ""
    return f"code={code}, msg={msg}, raw={json.dumps(payload, ensure_ascii=False)[:500]}"


def _generate_with_coze(context: dict) -> list[dict]:
    """调用扣子 v3 Chat：创建会话 -> 轮询状态 -> 获取answer消息 -> 解析TOP5。"""
    token = _get_cfg("COZE_API_KEY") or _get_cfg("COZE_TOKEN")
    bot_id = _get_cfg("COZE_BOT_ID")
    if not token or not bot_id:
        _set_status("扣子AI", False, "未调用：缺少 COZE_API_KEY/COZE_TOKEN 或 COZE_BOT_ID")
        return []

    try:
        import requests
    except Exception as e:
        _set_status("扣子AI", False, f"未调用：requests导入失败：{e}")
        return []

    base_url = (_get_cfg("COZE_BASE_URL") or "https://api.coze.cn").rstrip("/")
    user_id = _get_cfg("COZE_USER_ID", "audit_ppt_user")
    timeout_seconds = int(_get_cfg("COZE_TIMEOUT", "120"))
    poll_interval = float(_get_cfg("COZE_POLL_INTERVAL", "1.5"))

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    content = _system_prompt() + "\n\n" + _user_prompt(context)
    payload = {
        "bot_id": str(bot_id),
        "user_id": str(user_id),
        "stream": False,
        "auto_save_history": False,
        "additional_messages": [
            {"role": "user", "content_type": "text", "content": content}
        ],
    }

    try:
        _set_status("扣子AI", False, f"开始调用：base_url={base_url}, bot_id尾号={str(bot_id)[-6:]}")

        create_resp = requests.post(f"{base_url}/v3/chat", headers=headers, json=payload, timeout=30)
        if create_resp.status_code >= 400:
            _set_status("扣子AI", False, f"/v3/chat失败：HTTP {create_resp.status_code}，{create_resp.text[:500]}")
            return []

        create_json = create_resp.json()
        if create_json.get("code") not in (0, "0", None):
            _set_status("扣子AI", False, f"/v3/chat业务失败：{_coze_error_message(create_json)}")
            return []

        chat_data = create_json.get("data") or {}
        chat_id = chat_data.get("id") or chat_data.get("chat_id")
        conversation_id = chat_data.get("conversation_id")

        # 个别环境会直接同步返回answer，这里先尝试解析，避免漏掉。
        direct_rows = _safe_parse_json(json.dumps(create_json, ensure_ascii=False))
        if direct_rows:
            _set_status("扣子AI", True, f"同步返回解析成功，返回{len(direct_rows)}条可用结果")
            return direct_rows

        if not chat_id or not conversation_id:
            _set_status("扣子AI", False, f"返回缺少chat_id/conversation_id：{json.dumps(create_json, ensure_ascii=False)[:500]}")
            return []

        deadline = time.time() + timeout_seconds
        status = ""
        retrieve_json: dict[str, Any] = {}

        while time.time() < deadline:
            retrieve_resp = requests.get(
                f"{base_url}/v3/chat/retrieve",
                headers=headers,
                params={"conversation_id": conversation_id, "chat_id": chat_id},
                timeout=30,
            )
            if retrieve_resp.status_code >= 400:
                _set_status("扣子AI", False, f"retrieve失败：HTTP {retrieve_resp.status_code}，{retrieve_resp.text[:500]}")
                return []

            retrieve_json = retrieve_resp.json()
            if retrieve_json.get("code") not in (0, "0", None):
                _set_status("扣子AI", False, f"retrieve业务失败：{_coze_error_message(retrieve_json)}")
                return []

            status = (retrieve_json.get("data") or {}).get("status", "")
            if status == "completed":
                break
            if status in {"failed", "requires_action", "canceled"}:
                _set_status("扣子AI", False, f"运行未成功，status={status}，详情：{json.dumps(retrieve_json, ensure_ascii=False)[:500]}")
                return []
            time.sleep(poll_interval)

        if status != "completed":
            _set_status("扣子AI", False, f"调用超时，最后状态status={status}，详情：{json.dumps(retrieve_json, ensure_ascii=False)[:500]}")
            return []

        msg_resp = requests.get(
            f"{base_url}/v3/chat/message/list",
            headers=headers,
            params={"conversation_id": conversation_id, "chat_id": chat_id},
            timeout=30,
        )
        if msg_resp.status_code >= 400:
            _set_status("扣子AI", False, f"message/list失败：HTTP {msg_resp.status_code}，{msg_resp.text[:500]}")
            return []

        msg_json = msg_resp.json()
        if msg_json.get("code") not in (0, "0", None):
            _set_status("扣子AI", False, f"message/list业务失败：{_coze_error_message(msg_json)}")
            return []

        messages = _extract_messages_data(msg_json)
        candidate_texts: list[str] = []
        for msg in messages:
            role = msg.get("role")
            msg_type = msg.get("type")
            content_text = msg.get("content") or ""
            if not content_text:
                continue
            if role == "assistant" or msg_type in {"answer", "final_answer"}:
                candidate_texts.append(content_text)

        # 先从后往前解析，通常最后一条assistant/answer是最终答案。
        for content_text in reversed(candidate_texts):
            rows = _safe_parse_json(content_text)
            if rows:
                _set_status("扣子AI", True, f"调用成功，返回{len(rows)}条可用结果")
                return rows

        _set_status("扣子AI", False, f"已完成但未找到可解析answer，候选消息数={len(candidate_texts)}，返回前800字：{json.dumps(msg_json, ensure_ascii=False)[:800]}")
        return []

    except Exception as e:
        _set_status("扣子AI", False, f"调用异常：{type(e).__name__}: {e}")
        return []


def _get_ai_runtime() -> tuple[str, str, str]:
    api_key = _get_cfg("DINGTALK_API_KEY") or _get_cfg("DEAP_API_KEY") or _get_cfg("OPENAI_API_KEY") or _get_cfg("AI_API_KEY")
    base_url = _get_cfg("DINGTALK_BASE_URL") or _get_cfg("DEAP_BASE_URL") or _get_cfg("OPENAI_BASE_URL")
    model = _get_cfg("DINGTALK_MODEL") or _get_cfg("DEAP_MODEL") or _get_cfg("OPENAI_MODEL") or "gpt-4o-mini"
    return api_key, base_url, model


def _generate_with_openai_compatible(context: dict) -> list[dict]:
    api_key, base_url, model = _get_ai_runtime()
    if not api_key:
        _set_status("OpenAI兼容AI", False, "未调用：缺少API Key")
        return []
    try:
        from openai import OpenAI
    except Exception as e:
        _set_status("OpenAI兼容AI", False, f"openai导入失败：{e}")
        return []

    client = OpenAI(api_key=api_key, base_url=base_url) if base_url else OpenAI(api_key=api_key)
    try:
        kwargs = {
            "model": model,
            "messages": [
                {"role": "system", "content": _system_prompt()},
                {"role": "user", "content": _user_prompt(context)},
            ],
            "temperature": 0.15,
        }
        if model.startswith("gpt-4"):
            kwargs["response_format"] = {"type": "json_object"}
        _set_status("OpenAI兼容AI", False, f"开始调用：model={model}")
        resp = client.chat.completions.create(**kwargs)
        rows = _safe_parse_json(resp.choices[0].message.content or "")
        if rows:
            for row in rows:
                row["source"] = "AI智能总结"
            _set_status("OpenAI兼容AI", True, f"调用成功，返回{len(rows)}条可用结果")
            return rows
        return []
    except Exception as e:
        _set_status("OpenAI兼容AI", False, f"调用异常：{type(e).__name__}: {e}")
        return []


def generate_ai_top5(context: dict) -> list[dict]:
    coze_rows = _generate_with_coze(context)
    if coze_rows:
        for row in coze_rows:
            row["source"] = "AI智能总结（扣子）"
        return coze_rows

    rows = _generate_with_openai_compatible(context)
    if rows:
        return rows

    _set_status("AI总结", False, "AI总结未成功，已回退规则聚类")
    return []
