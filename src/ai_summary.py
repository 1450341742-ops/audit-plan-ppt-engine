from __future__ import annotations

import json
import os
from typing import Any


def _clean(value: Any) -> str:
    text = str(value or "").replace("\r", "\n")
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())


def _get_cfg(name: str, default: str = "") -> str:
    return os.getenv(name, default)


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


def _safe_parse_json(text: str) -> list[dict]:
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.replace("json\n", "", 1).strip()
    start = text.find("[")
    end = text.rfind("]")
    if start >= 0 and end > start:
        text = text[start : end + 1]
    data = json.loads(text)
    if not isinstance(data, list):
        return []
    rows = []
    for item in data[:5]:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "risk": _clean(item.get("高风险问题") or item.get("risk"))[:120],
                "analysis": _clean(item.get("风险维度分析") or item.get("analysis"))[:260],
                "advice": _clean(item.get("核查应对建议") or item.get("advice"))[:320],
                "score": 999,
            }
        )
    return rows


def _get_ai_runtime() -> tuple[str, str, str]:
    """Support OpenAI and DingTalk/DEAP OpenAI-compatible configs."""
    api_key = (
        _get_cfg("DINGTALK_API_KEY")
        or _get_cfg("DEAP_API_KEY")
        or _get_cfg("OPENAI_API_KEY")
        or _get_cfg("AI_API_KEY")
    )
    base_url = (
        _get_cfg("DINGTALK_BASE_URL")
        or _get_cfg("DEAP_BASE_URL")
        or _get_cfg("OPENAI_BASE_URL")
    )
    model = (
        _get_cfg("DINGTALK_MODEL")
        or _get_cfg("DEAP_MODEL")
        or _get_cfg("OPENAI_MODEL")
        or "gpt-4o-mini"
    )
    return api_key, base_url, model


def generate_ai_top5(context: dict) -> list[dict]:
    """Return AI-generated TOP5 risk rows. Empty list means fallback to rule summary."""
    api_key, base_url, model = _get_ai_runtime()
    if not api_key:
        return []

    try:
        from openai import OpenAI
    except Exception:
        return []

    client = OpenAI(api_key=api_key, base_url=base_url) if base_url else OpenAI(api_key=api_key)

    issues_text = _compact_issues(context)
    if not issues_text:
        return []

    system_prompt = """
你是临床试验第三方稽查与注册核查准备专家，擅长根据中心稽查发现，提炼最需要迎检准备的TOP5高风险问题。
你必须基于用户提供的问题内容总结，不要编造具体法规条款号，不要输出泛泛而谈。
输出必须是JSON数组，正好5项。每项必须包含：高风险问题、风险维度分析、核查应对建议。
要求：
1. 高风险问题：不是照搬原文，要概括成核查风险主题；可在一句中体现涉及的发现类型。
2. 风险维度分析：写清数据可靠性、受试者安全、方案依从性、受试者权益、伦理合规、药品/样本链条、疗效评价、注册核查等维度，说明为什么高风险。
3. 核查应对建议：按“立即行动/证据准备/系统改进或演练”写，必须可执行。
4. 5项之间不能重复，不能每项都写同样的话。
5. 每个字段控制在适合PPT表格展示的长度内。
6. 只输出JSON，不要输出解释。
""".strip()

    user_prompt = f"请根据以下中心稽查发现，生成TOP5高风险问题及核查应对建议。\n\n{issues_text}"

    try:
        kwargs = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.15,
        }
        if model.startswith("gpt-4"):
            kwargs["response_format"] = {"type": "json_object"}
        resp = client.chat.completions.create(**kwargs)
        content = resp.choices[0].message.content or ""
        try:
            obj = json.loads(content)
            if isinstance(obj, dict):
                data = obj.get("items") or obj.get("top5") or obj.get("data") or obj.get("results") or []
                return _safe_parse_json(json.dumps(data, ensure_ascii=False))
        except Exception:
            pass
        return _safe_parse_json(content)
    except Exception:
        return []
