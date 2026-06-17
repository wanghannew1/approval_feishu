import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
from dotenv import load_dotenv

from app.feishu_api import get_instance_detail, parse_form, extract_attachments
from app.batch_processor import get_approvers_with_roles, is_ready_for_print

load_dotenv()

st.set_page_config(page_title="详情页面", page_icon="📋", layout="wide")

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

STATUS_DISPLAY = {
    "PENDING": "审批中",
    "APPROVED": "已通过",
    "REJECTED": "已拒绝",
    "CANCELED": "已撤销",
    "RUNNING": "审批中",
}

form_widgets = parse_form(detail)
form_title = ""
for widget in form_widgets:
    if widget.get("name") == "标题":
        form_title = widget.get("value", "")
        break
title = form_title or detail.get("approval_name", "无标题")
raw_status = detail.get("status", "")
status_text = STATUS_DISPLAY.get(raw_status, raw_status)
ready = is_ready_for_print(detail)
if ready and raw_status == "RUNNING":
    status_text = "审批完成待出纳办理"

st.title(title)
st.caption(f"状态: {status_text}  |  单号: {code}")

st.divider()

tab1, tab2, tab3 = st.tabs(["表单字段", "审批人", "附件"])

form_widgets = parse_form(detail)

with tab1:
    non_attachment = [w for w in form_widgets if w.get("type") != "attachmentV2"]
    if non_attachment:
        form_data = []
        for w in non_attachment:
            form_data.append({
                "字段": w.get("name", ""),
                "类型": w.get("type", ""),
                "值": str(w.get("value", "")) if w.get("value") else "",
            })
        st.dataframe(form_data, width="stretch", hide_index=True)
    else:
        st.info("无文本表单字段")

with tab2:
    approvers = get_approvers_with_roles(detail)
    if approvers:
        approver_data = []
        for a in approvers:
            a_status = STATUS_DISPLAY.get(a.get("status", ""), a.get("status", ""))
            approver_data.append({
                "审批人": a.get("approver_name", ""),
                "角色": a.get("role") or "—",
                "状态": a_status,
            })
        st.dataframe(approver_data, width="stretch", hide_index=True)
    else:
        st.info("无审批人信息")

with tab3:
    attachments = extract_attachments(form_widgets)
    if attachments:
        for att in attachments:
            field_name = att.get("field_name", "附件")
            values = att.get("value", [])
            st.markdown(f"**{field_name}** ({len(values)} 个文件)")
            for v in values:
                st.code(v, language=None)
    else:
        st.info("无附件")
