from __future__ import annotations

import json
import os
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
    if not isinstance(data, list):
        return []
    rows = []
    for item in data[:5]:
        if not isinstance(item, dict):
            continue
        risk = _clean(item.get("高风险问题") or item.get("risk") or item.get("问题") or item.get("title"))
        analysis = _clean(item.get("风险维度分析") or item.get("analysis") or item.get("风险分析"))
        advice = _clean(item.get("核查应对建议") or item.get("advice") or item.get("建议"))
        if not risk or not analysis or not advice:
            continue
        rows.append({"risk": risk[:120], "analysis": analysis[:260], "advice": advice[:320], "score": 999, "source": "AI智能总结"})
    return rows


def _safe_parse_json(text: str) -> list[dict]:
    text = (text or "").strip()
    if not text:
        _set_status("AI解析", False, "AI返回为空")
        return []
    if text.startswith("```"):
        text = text.strip("`")
        text = text.replace("json\n", "", 1).strip()
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            data = obj.get("items") or obj.get("top5") or obj.get("data") or obj.get("results") or obj.get("list") or []
            if isinstance(data, list):
                obj = data
        rows = _normalize_rows(obj)
        if rows:
            return rows
        _set_status("AI解析", False, f"AI返回JSON但字段不匹配或条数不足，前200字：{text[:200]}")
        return []
    except Exception:
        start = text.find("[")
        end = text.rfind("]")
        if start >= 0 and end > start:
            try:
                rows = _normalize_rows(json.loads(text[start : end + 1]))
                if rows:
                    return rows
            except Exception as e:
                _set_status("AI解析", False, f"JSON数组截取解析失败：{e}")
                return []
        _set_status("AI解析", False, f"JSON解析失败，AI返回前200字：{text[:200]}")
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


def _generate_with_coze(context: dict) -> list[dict]:
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
    timeout_seconds = int(_get_cfg("COZE_TIMEOUT", "90"))

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    content = _system_prompt() + "\n\n" + _user_prompt(context)
    payload = {
        "bot_id": bot_id,
        "user_id": user_id,
        "stream": False,
        "auto_save_history": False,
        "additional_messages": [{"role": "user", "content_type": "text", "content": content}],
    }

    try:
        _set_status("扣子AI", False, f"开始调用：base_url={base_url}, bot_id尾号={bot_id[-6:]}")
        resp = requests.post(f"{base_url}/v3/chat", headers=headers, json=payload, timeout=timeout_seconds)
        if resp.status_code >= 400:
            _set_status("扣子AI", False, f"/v3/chat失败：HTTP {resp.status_code}，{resp.text[:300]}")
            return []
        data = resp.json()
        chat_id = (data.get("data") or {}).get("id")
        conversation_id = (data.get("data") or {}).get("conversation_id")
        if not chat_id or not conversation_id:
            _set_status("扣子AI", False, f"返回缺少chat_id/conversation_id：{json.dumps(data, ensure_ascii=False)[:300]}")
            return _safe_parse_json(json.dumps(data, ensure_ascii=False))

        deadline = time.time() + timeout_seconds
        status = ""
        while time.time() < deadline:
            check = requests.get(
                f"{base_url}/v3/chat/retrieve",
                headers=headers,
                params={"chat_id": chat_id, "conversation_id": conversation_id},
                timeout=20,
            )
            if check.status_code >= 400:
                _set_status("扣子AI", False, f"retrieve失败：HTTP {check.status_code}，{check.text[:300]}")
                return []
            status = (check.json().get("data") or {}).get("status", "")
            if status in {"completed", "failed", "requires_action", "canceled"}:
                break
            time.sleep(1.2)
        if status != "completed":
            _set_status("扣子AI", False, f"运行未完成或失败，status={status}")
            return []

        msg_resp = requests.get(
            f"{base_url}/v3/chat/message/list",
            headers=headers,
            params={"chat_id": chat_id, "conversation_id": conversation_id},
            timeout=20,
        )
        if msg_resp.status_code >= 400:
            _set_status("扣子AI", False, f"message/list失败：HTTP {msg_resp.status_code}，{msg_resp.text[:300]}")
            return []
        msg_json = msg_resp.json()
        messages = _extract_messages_data(msg_json)
        for msg in messages:
            if msg.get("type") == "answer" and msg.get("content"):
                rows = _safe_parse_json(msg.get("content", ""))
                if rows:
                    _set_status("扣子AI", True, f"调用成功，返回{len(rows)}条可用结果")
                    return rows
        _set_status("扣子AI", False, f"未找到可解析answer消息，返回前500字：{json.dumps(msg_json, ensure_ascii=False)[:500]}")
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
            "messages": [{"role": "system", "content": _system_prompt()}, {"role": "user", "content": _user_prompt(context)}],
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
