import json
import sys
from datetime import datetime
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
from dotenv import load_dotenv

from app.feishu_api import extract_attachments, parse_form

load_dotenv()

USER_MAPPING_PATH = _PROJECT_ROOT / "user_mapping.json"

STATUS_BADGE = {
    "PENDING": "🟡 审批中",
    "APPROVED": "🟢 已通过",
    "REJECTED": "🔴 已拒绝",
    "CANCELED": "⚪ 已撤销",
    "RUNNING": "🟡 审批中",
}
STATUS_LABEL = {
    "PENDING": "审批中",
    "APPROVED": "已通过",
    "REJECTED": "已拒绝",
    "CANCELED": "已撤销",
    "RUNNING": "审批中",
}


def _load_user_mapping():
    try:
        with open(USER_MAPPING_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _resolve_name(uid, mapping):
    if not uid:
        return ""
    return mapping.get(uid, uid)


def _fmt(ts_str):
    try:
        ts = int(ts_str) / 1000
        if ts <= 0:
            return ""
        return datetime.fromtimestamp(ts).strftime("%m月%d日 %H:%M")
    except (ValueError, TypeError):
        return ""


def _waiting_time(ts_str):
    try:
        start = int(ts_str) / 1000
        if start <= 0:
            return ""
        delta = datetime.now().timestamp() - start
        days = int(delta / 86400)
        if days >= 1:
            return f"已等待 {days} 天"
        hours = int(delta / 3600)
        return f"已等待 {hours} 小时"
    except (ValueError, TypeError):
        return ""


# ── page setup ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="审批详情", page_icon="📋", layout="wide")

code = st.session_state.get("detail_code")
if not code:
    st.warning("未选择审批单")
    if st.button("← 返回列表"):
        st.switch_page("app.py")
    st.stop()

if st.button("← 返回列表", key="back"):
    st.session_state.detail_code = None
    st.switch_page("app.py")

detail = st.session_state.instance_details_cache.get(code)
if not detail:
    st.spinner("加载中...")
    st.stop()

# ── data ─────────────────────────────────────────────────────────────────────
form_widgets = parse_form(detail)
user_mapping = _load_user_mapping()

serial = detail.get("serial_number") or code
approval_name = detail.get("approval_name", "")
raw_status = detail.get("status", "")
start_time = _fmt(str(detail.get("start_time", "")))
submitter_name = _resolve_name(detail.get("user_id", ""), user_mapping) or detail.get("user_id", "")

# ── header ───────────────────────────────────────────────────────────────────
st.caption(f"编号：{serial}")
st.subheader(approval_name)

col_status, col_meta = st.columns([1, 3])
with col_status:
    badge = STATUS_BADGE.get(raw_status, raw_status)
    st.markdown(f"### {badge}")
with col_meta:
    if submitter_name:
        st.caption(f"{submitter_name} 提交于 {start_time}")

st.divider()

# ── 审批详情 ──
st.subheader("审批详情")
for w in form_widgets:
    w_type = w.get("type", "")
    w_name = w.get("name", "")
    w_value = w.get("value", "")

    if w_type == "attachmentV2":
        continue

    if w_type == "fieldList":
        with st.expander(f"📋 {w_name}", expanded=True):
            if isinstance(w_value, list):
                for row_item in w_value:
                    if isinstance(row_item, list):
                        for item in row_item:
                            sub_name = item.get("name", "")
                            sub_val = item.get("value", "")
                            sub_type = item.get("type", "")
                            if sub_type == "amount":
                                ext = item.get("ext", {})
                                capital = ext.get("capitalValue", "")
                                cur = ext.get("currency", "CNY")
                                val_fmt = f"{sub_val:,.2f}" if isinstance(sub_val, (int, float)) else str(sub_val)
                                st.write(f"**{sub_name}**  {val_fmt} {cur}-人民币元")
                                if capital:
                                    st.caption(capital)
                            else:
                                st.write(f"**{sub_name}**  {sub_val}")
                    else:
                        st.write(str(row_item))
        continue

    st.markdown(f"**{w_name}**")
    st.write(str(w_value))
    st.divider()

attachments = extract_attachments(form_widgets)
if attachments:
    for att in attachments:
        field_name = att.get("field_name", "附件")
        vals = att.get("value", [])
        st.markdown(f"📎 **{field_name}**")
        for v in vals:
            fname = att.get("ext", "") or v.rsplit("/", 1)[-1].split("?")[0] if v else "文件"
            st.markdown(f"📥 [{fname}]({v})")

# ── 审批记录 ──
st.divider()

st.subheader("审批记录")

records = []

for event in detail.get("timeline", []):
    if event.get("type") == "START":
        records.append({
            "节点名称": "提交",
            "审批人": _resolve_name(event.get("user_id", ""), user_mapping) or event.get("user_id", ""),
            "审批结果": "已提交",
            "审批意见": "",
            "审批时间": _fmt(str(event.get("create_time", ""))),
        })

for task in detail.get("task_list", []):
    t_status = task.get("status", "")
    result = "审批中" if t_status == "PENDING" else STATUS_LABEL.get(t_status, t_status)
    task_time = _waiting_time(str(task.get("start_time", ""))) if t_status == "PENDING" else _fmt(str(task.get("start_time", "")))
    records.append({
        "节点名称": task.get("node_name", ""),
        "审批人": _resolve_name(task.get("user_id", ""), user_mapping) or task.get("user_id", ""),
        "审批结果": result,
        "审批意见": "",
        "审批时间": task_time,
    })

for a in detail.get("approver_list", []):
    records.append({
        "节点名称": "",
        "审批人": a.get("approver_name", ""),
        "审批结果": STATUS_LABEL.get(a.get("status", ""), a.get("status", "")),
        "审批意见": a.get("comment", ""),
        "审批时间": _fmt(str(a.get("approval_time", ""))),
    })

end_time_raw = detail.get("end_time", "")
if end_time_raw and end_time_raw != "0":
    records.append({
        "节点名称": "结束",
        "审批人": "系统",
        "审批结果": STATUS_LABEL.get(raw_status, raw_status),
        "审批意见": "",
        "审批时间": _fmt(str(end_time_raw)),
    })
else:
    records.append({
        "节点名称": "结束",
        "审批人": "系统",
        "审批结果": "未结束",
        "审批意见": "",
        "审批时间": "",
    })

if records:
    st.dataframe(
        records,
        width="stretch",
        hide_index=True,
        column_config={
            "节点名称": st.column_config.TextColumn("节点名称", width="small"),
            "审批人": st.column_config.TextColumn("审批人", width="small"),
            "审批结果": st.column_config.TextColumn("审批结果", width="small"),
            "审批意见": st.column_config.TextColumn("审批意见", width="medium"),
            "审批时间": st.column_config.TextColumn("审批时间", width="medium"),
        },
    )
else:
    st.info("无审批记录")
