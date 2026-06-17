"""
Streamlit application entry point.

Provides a web UI for querying and downloading approval attachments,
inserting signatures, and printing payroll sheets.
"""

import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
from dotenv import load_dotenv

from app.batch_processor import (
    get_approvers_with_roles,
    is_approval_passed,
    is_ready_for_print,
    process_single_approval,
)
from app.feishu_api import (
    download_file,
    extract_attachments,
    get_definition,
    get_instance_detail,
    get_tenant_token,
    list_instances,
    parse_form,
)

load_dotenv()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFINITIONS_FILE = Path(__file__).parent / "approval_definitions.json"
SETTINGS_FILE = Path(__file__).parent / "settings.json"

_DEFAULT_DEFINITIONS = {
    "1CF34ABB-781C-40B0-9A4F-3CC416612423": "项目人员工资发放审批单（系统工资单）",
}


def _load_definitions():
    try:
        with open(DEFINITIONS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return dict(_DEFAULT_DEFINITIONS)


def _save_definitions(data):
    with open(DEFINITIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _get_credentials():
    settings = _load_settings()
    return (
        settings.get("app_id") or os.getenv("FEISHU_APP_ID", ""),
        settings.get("app_secret") or os.getenv("FEISHU_APP_SECRET", ""),
    )

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="飞书审批打印工具",
    page_icon="🖨️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
if "token" not in st.session_state:
    st.session_state.token = None
if "query_results" not in st.session_state:
    st.session_state.query_results = []
if "selected_instances" not in st.session_state:
    st.session_state.selected_instances = set()
if "instance_details_cache" not in st.session_state:
    st.session_state.instance_details_cache = {}
if "selected_definition" not in st.session_state:
    # Default to the payroll sheet approval
    st.session_state.selected_definition = "1CF34ABB-781C-40B0-9A4F-3CC416612423"

STATUS_LABELS = {
    "全部": None,
    "审批完成待出纳办理": "READY_FOR_PRINT",
    "审批中": "PENDING",
    "已通过": "APPROVED",
    "已拒绝": "REJECTED",
}

STATUS_DISPLAY = {
    "PENDING": "审批中",
    "APPROVED": "已通过",
    "REJECTED": "已拒绝",
    "CANCELED": "已撤销",
    "RUNNING": "审批中",
}


def _get_token(app_id: str, app_secret: str) -> str:
    """Get or refresh tenant access token."""
    if st.session_state.token is None:
        try:
            st.session_state.token = get_tenant_token(app_id, app_secret)
        except RuntimeError as e:
            st.error(f"获取 Token 失败: {e}")
            return ""
    return st.session_state.token


def _ms_timestamp(dt: datetime) -> str:
    """Convert datetime to millisecond timestamp string for Feishu API."""
    return str(int(dt.timestamp() * 1000))


def _format_datetime(ts_str: str) -> str:
    """Format a millisecond timestamp string to readable datetime."""
    try:
        ts = int(ts_str) / 1000
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
    except (ValueError, TypeError):
        return ts_str


def render_sidebar():
    definitions = _load_definitions()

    with st.sidebar:
        st.header("审批单分类")

        if not definitions:
            st.info("暂无模板，请在下方新增")

        for code, name in definitions.items():
            selected = st.session_state.selected_definition == code
            cols = st.columns([8, 1, 1])
            with cols[0]:
                if st.button(
                    name,
                    key=f"def_{code}",
                    use_container_width=True,
                    type="primary" if selected else "secondary",
                ):
                    st.session_state.selected_definition = code
                    st.rerun()
            with cols[1]:
                with st.popover("✏️"):
                    new_name = st.text_input("模板名称", value=name, key=f"edit_name_{code}")
                    if st.button("保存", key=f"save_{code}"):
                        definitions[code] = new_name.strip()
                        _save_definitions(definitions)
                        st.rerun()
            with cols[2]:
                if st.button("🗑️", key=f"del_{code}"):
                    del definitions[code]
                    if st.session_state.selected_definition == code:
                        st.session_state.selected_definition = next(iter(definitions), "")
                    _save_definitions(definitions)
                    st.rerun()

        st.divider()

        with st.expander("➕ 新增审批模板"):
            new_code = st.text_input("definitionCode", key="new_code")
            fetch_clicked = st.button("🔍 获取模板名称", key="fetch_name", use_container_width=True)

            if fetch_clicked and new_code.strip():
                app_id, app_secret = _get_credentials()
                if not app_id or not app_secret:
                    st.error("请先在设置中配置 App ID 和 App Secret")
                else:
                    token = _get_token(app_id, app_secret)
                    if token:
                        try:
                            info = get_definition(token, new_code.strip())
                            st.session_state.new_fetched_name = info.get("approval_name", "")
                        except RuntimeError as e:
                            st.error(f"获取失败: {e}")
                            st.session_state.new_fetched_name = ""

            fetched_name = st.session_state.get("new_fetched_name", "")
            new_name = st.text_input("模板名称", value=fetched_name, key="new_name")
            if st.button("添加", key="add_def", use_container_width=True):
                code = new_code.strip()
                name = new_name.strip()
                if code and name and code not in definitions:
                    definitions[code] = name
                    _save_definitions(definitions)
                    st.session_state.selected_definition = code
                    st.session_state.new_fetched_name = ""
                    st.rerun()
                elif code in definitions:
                    st.error("该 definitionCode 已存在")

        st.divider()

        with st.expander("⚙️ 设置"):
            settings = _load_settings()
            app_id_val = st.text_input(
                "App ID",
                value=settings.get("app_id", os.getenv("FEISHU_APP_ID", "")),
                type="password",
                key="settings_app_id",
            )
            app_secret_val = st.text_input(
                "App Secret",
                value=settings.get("app_secret", os.getenv("FEISHU_APP_SECRET", "")),
                type="password",
                key="settings_app_secret",
            )
            if st.button("保存设置", key="save_settings", use_container_width=True):
                settings["app_id"] = app_id_val.strip()
                settings["app_secret"] = app_secret_val.strip()
                _save_settings(settings)
                st.session_state.token = None
                st.success("已保存")

    if st.session_state.selected_definition not in definitions and definitions:
        st.session_state.selected_definition = next(iter(definitions))

    return st.session_state.selected_definition


def render_query_panel(approval_code):
    if not approval_code:
        st.warning("请在左侧添加审批模板")
        return
    st.header("📋 查询审批实例")

    col_status, col_start, col_end = st.columns([2, 2, 2])

    with col_status:
        status_option = st.selectbox(
            "审批状态",
            options=list(STATUS_LABELS.keys()),
            index=1,  # Default: "审批完成待出纳办理"
            key="status_filter",
        )

    with col_start:
        start_date = st.date_input(
            "开始日期",
            value=datetime.now() - timedelta(days=30),
            key="start_date",
        )

    with col_end:
        end_date = st.date_input(
            "结束日期",
            value=datetime.now(),
            key="end_date",
        )

    if st.button("🔍 查询", use_container_width=True, type="primary"):
        app_id, app_secret = _get_credentials()

        if not app_id or not app_secret:
            st.error("请填写 App ID 和 App Secret")
            return

        if not approval_code:
            st.error("请填写审批定义 Code")
            return

        token = _get_token(app_id, app_secret)
        if not token:
            return

        with st.spinner("正在查询..."):
            try:
                start_ms = _ms_timestamp(datetime.combine(start_date, datetime.min.time()))
                end_ms = _ms_timestamp(
                    datetime.combine(end_date, datetime.max.time())
                )

                # First, get all instance codes in the date range
                instance_codes = list_instances(
                    token, approval_code, start_ms, end_ms
                )

                if not instance_codes:
                    st.warning("未找到审批实例")
                    st.session_state.query_results = []
                    return

                # Then get details for each instance
                results = []
                progress = st.progress(0, text="获取实例详情...")

                for i, code in enumerate(instance_codes):
                    progress.progress(
                        (i + 1) / len(instance_codes),
                        text=f"获取详情 {i + 1}/{len(instance_codes)}",
                    )
                    try:
                        detail = get_instance_detail(token, code)
                        # Cache the detail
                        st.session_state.instance_details_cache[code] = detail
                        results.append(detail)
                    except RuntimeError as e:
                        st.warning(f"获取实例 {code[:8]}... 详情失败: {e}")

                progress.empty()

                # Apply status filter
                status_val = STATUS_LABELS[status_option]
                if status_val == "READY_FOR_PRINT":
                    # Special filter: RUNNING instances where all mandatory roles approved
                    filtered = [
                        r for r in results if is_ready_for_print(r)
                    ]
                elif status_val is not None:
                    filtered = [
                        r for r in results
                        if r.get("status") == status_val
                    ]
                else:
                    filtered = results

                st.session_state.query_results = filtered
                st.session_state.selected_instances = set()

                if not filtered:
                    st.info(
                        f"共 {len(results)} 个实例，筛选后无匹配结果"
                    )
                else:
                    st.success(
                        f"共 {len(results)} 个实例，筛选后 {len(filtered)} 个"
                    )

            except RuntimeError as e:
                st.error(f"查询失败: {e}")


def render_instance_list() -> None:
    """Render the instance list with checkboxes for batch selection."""
    results = st.session_state.query_results
    if not results:
        return

    st.header("📝 审批实例列表")

    # Select all checkbox
    select_all = st.checkbox("全选", key="select_all")

    # Build display data
    display_rows = []
    for idx, detail in enumerate(results):
        code = detail.get("instance_code", "")
        title = detail.get("title", "无标题")
        raw_status = detail.get("status", "")
        status_text = STATUS_DISPLAY.get(raw_status, raw_status)
        create_time = detail.get("start_time", "")
        if create_time:
            create_time = _format_datetime(str(create_time))

        # Determine if ready for print
        ready = is_ready_for_print(detail)
        status_label = status_text
        if ready and raw_status == "RUNNING":
            status_label = "审批完成待出纳办理"

        display_rows.append({
            "序号": idx + 1,
            "审批单号": code,
            "标题": title,
            "状态": status_label,
            "提交时间": create_time,
        })

    # Display as dataframe
    if display_rows:
        st.dataframe(
            display_rows,
            use_container_width=True,
            hide_index=True,
            column_config={
                "序号": st.column_config.NumberColumn(width="small"),
                "审批单号": st.column_config.TextColumn(width="medium"),
                "标题": st.column_config.TextColumn(width="large"),
                "状态": st.column_config.TextColumn(width="small"),
                "提交时间": st.column_config.TextColumn(width="small"),
            },
        )

    # Individual selection
    st.subheader("选择实例")
    selected = set()

    for idx, detail in enumerate(results):
        code = detail.get("instance_code", "")
        title = detail.get("title", "无标题")
        raw_status = detail.get("status", "")
        status_text = STATUS_DISPLAY.get(raw_status, raw_status)
        ready = is_ready_for_print(detail)
        if ready and raw_status == "RUNNING":
            status_text = "审批完成待出纳办理"

        label = f"**{idx + 1}.** {title} — {status_text}"

        if select_all:
            default = True
        else:
            default = code in st.session_state.selected_instances

        checked = st.checkbox(label, value=default, key=f"sel_{code}")
        if checked:
            selected.add(code)

    st.session_state.selected_instances = selected


def render_batch_actions():
    selected = st.session_state.selected_instances
    if not selected:
        return

    st.header("🔧 批量操作")
    st.info(f"已选择 {len(selected)} 个实例")

    col_download, col_print = st.columns(2)

    with col_download:
        if st.button("📥 下载附件", use_container_width=True):
            _handle_download()

    with col_print:
        if st.button("✍️ 签名并打印", use_container_width=True):
            _handle_sign_and_print()


def _handle_download():
    app_id, app_secret = _get_credentials()
    token = _get_token(app_id, app_secret)
    if not token:
        return

    results = st.session_state.query_results
    selected = st.session_state.selected_instances

    save_dir = Path("./downloads")
    save_dir.mkdir(parents=True, exist_ok=True)

    total = len(selected)
    progress = st.progress(0, text=f"下载附件 0/{total}")
    downloaded_count = 0

    for i, detail in enumerate(results):
        code = detail.get("instance_code", "")
        if code not in selected:
            continue

        try:
            form_widgets = parse_form(detail)
            attachments = extract_attachments(form_widgets)

            if not attachments:
                progress.progress(
                    (i + 1) / total,
                    text=f"下载附件 {i + 1}/{total} — 无附件",
                )
                continue

            instance_dir = save_dir / code[:12]
            instance_dir.mkdir(parents=True, exist_ok=True)

            for att in attachments:
                field_name = att.get("field_name", "附件")
                values = att.get("value", [])
                for val in values:
                    try:
                        filepath = download_file(token, val, str(instance_dir))
                        downloaded_count += 1
                        progress.progress(
                            (i + 1) / total,
                            text=f"下载附件 {i + 1}/{total} — {Path(filepath).name}",
                        )
                    except RuntimeError as e:
                        st.warning(f"下载 {field_name} 失败: {e}")

        except Exception as e:
            st.warning(f"处理实例 {code[:8]}... 失败: {e}")

    progress.empty()
    st.success(f"下载完成，共下载 {downloaded_count} 个文件")


def _handle_sign_and_print():
    app_id, app_secret = _get_credentials()
    token = _get_token(app_id, app_secret)
    if not token:
        return

    selected = st.session_state.selected_instances
    total = len(selected)
    progress = st.progress(0, text=f"签名并打印 0/{total}")

    success_count = 0
    for i, code in enumerate(selected):
        try:
            result = process_single_approval(
                code,
                token,
                {
                    "save_dir": "./downloads",
                    "signatures_dir": "./signatures",
                },
            )
            if result["success"]:
                success_count += 1
                progress.progress(
                    (i + 1) / total,
                    text=f"✅ {result['title'][:20]}: {result['message']}",
                )
            elif result.get("skipped"):
                progress.progress(
                    (i + 1) / total,
                    text=f"⏭️ {result['title'][:20]}: {result['message']}",
                )
            else:
                progress.progress(
                    (i + 1) / total,
                    text=f"❌ {result['title'][:20]}: {result['message']}",
                )
        except Exception as e:
            progress.progress(
                (i + 1) / total,
                text=f"❌ {code[:8]}...: {e}",
            )

    progress.empty()
    st.success(f"处理完成: {success_count}/{total} 成功")


def render_instance_detail() -> None:
    """Render expandable instance detail panels."""
    results = st.session_state.query_results
    if not results:
        return

    st.header("📄 实例详情")

    for detail in results:
        code = detail.get("instance_code", "")
        title = detail.get("title", "无标题")
        raw_status = detail.get("status", "")
        status_text = STATUS_DISPLAY.get(raw_status, raw_status)
        ready = is_ready_for_print(detail)
        if ready and raw_status == "RUNNING":
            status_text = "审批完成待出纳办理"

        with st.expander(f"{title} ({status_text})"):
            # Form fields
            st.subheader("表单字段")
            form_widgets = parse_form(detail)
            if form_widgets:
                form_data = []
                for w in form_widgets:
                    wtype = w.get("type", "")
                    wname = w.get("name", "")
                    wvalue = w.get("value", "")
                    # Skip attachmentV2 — shown separately
                    if wtype == "attachmentV2":
                        continue
                    form_data.append({
                        "字段": wname,
                        "类型": wtype,
                        "值": str(wvalue) if wvalue else "",
                    })
                if form_data:
                    st.dataframe(form_data, use_container_width=True, hide_index=True)
                else:
                    st.info("无文本表单字段")
            else:
                st.info("无表单数据")

            # Approver list
            st.subheader("审批人列表")
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
                st.dataframe(approver_data, use_container_width=True, hide_index=True)
            else:
                st.info("无审批人信息")

            # Attachments
            st.subheader("附件")
            attachments = extract_attachments(form_widgets)
            if attachments:
                for att in attachments:
                    field_name = att.get("field_name", "附件")
                    values = att.get("value", [])
                    st.markdown(f"**{field_name}**: {len(values)} 个文件")
                    for v in values:
                        st.code(v, language=None)
            else:
                st.info("无附件")


def main():
    approval_code = render_sidebar()
    render_query_panel(approval_code)
    render_instance_list()
    render_batch_actions()
    render_instance_detail()


if __name__ == "__main__":
    main()