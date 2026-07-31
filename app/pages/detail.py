import sys
from datetime import datetime
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
from dotenv import load_dotenv

from app import config_store
from app.feishu_api import extract_attachments, parse_form

load_dotenv()

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
    return config_store.load_or_create(
        config_store.PATH_USER_MAPPING, config_store.DEFAULT_USER_MAPPING
    )


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


def _load_role_mapping_file():
    return config_store.load_or_create(
        config_store.PATH_ROLE_MAPPING, config_store.DEFAULT_ROLE_MAPPING
    )


def _load_workflow_order():
    order = config_store.load_or_create(
        config_store.PATH_WORKFLOW_ORDER, config_store.DEFAULT_WORKFLOW_ORDER
    )
    return {k: v for k, v in order.items() if not k.startswith("_")}


def _get_sort_key(node_name):
    role_mapping = _load_role_mapping_file()
    display_name = role_mapping.get(node_name, node_name)
    order = _load_workflow_order()
    return order.get(node_name, order.get(display_name, 999))


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
st.caption(f"编号：{serial}  |  instance_id: `{code}`")
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
            if not isinstance(w_value, list) or not w_value:
                continue

            id_to_name = {}
            all_columns = []
            rows_data = []
            for row_item in w_value:
                if not isinstance(row_item, list):
                    continue
                row_dict = {}
                for item in row_item:
                    name = item.get("name", "")
                    val = item.get("value", "")
                    itype = item.get("type", "")
                    item_id = item.get("id", "")
                    if item_id and item_id not in id_to_name:
                        id_to_name[item_id] = name
                        all_columns.append(name)
                    if itype == "amount" and isinstance(val, (int, float)):
                        row_dict[name] = f"{val:,.2f}"
                    else:
                        row_dict[name] = str(val) if val else ""
                rows_data.append(row_dict)

            col_names = [c for c in all_columns if c in rows_data[0]] if rows_data else []
            table_data = [{c: rd.get(c, "") for c in col_names} for rd in rows_data]

            ext_items = w.get("ext")
            if isinstance(ext_items, list) and ext_items:
                summary = {}
                for ei in ext_items:
                    eid = ei.get("id", "")
                    name = id_to_name.get(eid, eid)
                    val = ei.get("value", "")
                    if ei.get("type") == "amount" and val:
                        try:
                            summary[name] = f"{float(val):,.2f}"
                        except ValueError:
                            summary[name] = val
                    elif val:
                        summary[name] = str(val)
                if summary:
                    summary_row = {}
                    for c in col_names:
                        if c in summary:
                            summary_row[c] = summary[c]
                        elif c == col_names[0]:
                            summary_row[c] = "**汇总**"
                        else:
                            summary_row[c] = ""
                    table_data.append(summary_row)

            st.dataframe(table_data, width="stretch", hide_index=True)
        continue

    st.markdown(f"**{w_name}**")
    st.write(str(w_value))
    st.divider()

attachments = extract_attachments(form_widgets)
if attachments:
    for att in attachments:
        field_name = att.get("field_name", "附件")
        vals = att.get("value", [])
        ext_str = att.get("ext", "")
        fnames = [f.strip() for f in ext_str.split(",") if f.strip()] if ext_str else []
        st.markdown(f"📎 **{field_name}**")
        for i, v in enumerate(vals):
            fname = fnames[i] if i < len(fnames) else (v.rsplit("/", 1)[-1].split("?")[0] if v else "文件")
            st.markdown(f"📥 [{fname}]({v})")

# ── 审批记录 ──
st.divider()

st.subheader("审批记录")

records = []
workflow_order = _load_workflow_order()

for event in detail.get("timeline", []):
    if event.get("type") == "START":
        records.append({
            "节点名称": "提交",
            "审批人": _resolve_name(event.get("user_id", ""), user_mapping) or event.get("user_id", ""),
            "审批结果": "已提交",
            "审批意见": "",
            "审批时间": _fmt(str(event.get("create_time", ""))),
            "_order": workflow_order.get("提交", 0),
        })

for task in detail.get("task_list", []):
    t_status = task.get("status", "")
    node = task.get("node_name", "")
    result = "审批中" if t_status == "PENDING" else STATUS_LABEL.get(t_status, t_status)
    task_time = _waiting_time(str(task.get("start_time", ""))) if t_status == "PENDING" else _fmt(str(task.get("start_time", "")))
    records.append({
        "节点名称": node,
        "审批人": _resolve_name(task.get("user_id", ""), user_mapping) or task.get("user_id", ""),
        "审批结果": result,
        "审批意见": "",
        "审批时间": task_time,
        "_order": _get_sort_key(node),
    })

for a in detail.get("approver_list", []):
    records.append({
        "节点名称": "",
        "审批人": a.get("approver_name", ""),
        "审批结果": STATUS_LABEL.get(a.get("status", ""), a.get("status", "")),
        "审批意见": a.get("comment", ""),
        "审批时间": _fmt(str(a.get("approval_time", ""))),
        "_order": 999,
    })

end_time_raw = detail.get("end_time", "")
if end_time_raw and end_time_raw != "0":
    records.append({
        "节点名称": "结束",
        "审批人": "系统",
        "审批结果": STATUS_LABEL.get(raw_status, raw_status),
        "审批意见": "",
        "审批时间": _fmt(str(end_time_raw)),
        "_order": workflow_order.get("结束", 999),
    })
else:
    records.append({
        "节点名称": "结束",
        "审批人": "系统",
        "审批结果": "未结束",
        "审批意见": "",
        "审批时间": "",
        "_order": workflow_order.get("结束", 999),
    })

records.sort(key=lambda r: r.pop("_order", 999))

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
