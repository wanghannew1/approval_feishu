import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
import json

from app.payroll_merger import (
    DEFAULT_MERGE_CONFIG,
    batch_print,
    check_wps_available,
    get_last_wps_error,
    _load_mapping_rules,
    merge_payrolls_simple,
)

# ── settings persistence ─────────────────────────────────────────────────────
_SETTINGS_FILE = _PROJECT_ROOT / "settings.json"

_MERGE_CFG_KEY = "merge_layout_config"

def _load_merge_settings() -> dict:
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "payroll_dir": data.get("merge_payroll_dir", ""),
            "output_dir": data.get("merge_output_dir", ""),
        }
    except (FileNotFoundError, json.JSONDecodeError):
        return {"payroll_dir": "", "output_dir": ""}

def _save_merge_settings(payroll_dir: str, output_dir: str) -> None:
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    data["merge_payroll_dir"] = payroll_dir
    data["merge_output_dir"] = output_dir
    with open(_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _load_merge_layout_config() -> dict:
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get(_MERGE_CFG_KEY, {})
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def _save_merge_layout_config(overrides: dict) -> None:
    try:
        with open(_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}
    data[_MERGE_CFG_KEY] = overrides
    with open(_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ── page setup ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="工资表合并", page_icon="🗂️", layout="wide")

# ── session state ────────────────────────────────────────────────────────────
if "merge_results" not in st.session_state:
    st.session_state.merge_results = None
if "wps_available" not in st.session_state:
    st.session_state.wps_available = check_wps_available()

_saved = _load_merge_settings()
if "merge_payroll_dir" not in st.session_state or not st.session_state.merge_payroll_dir.strip():
    st.session_state.merge_payroll_dir = _saved["payroll_dir"]
if "merge_output_dir" not in st.session_state or not st.session_state.merge_output_dir.strip():
    st.session_state.merge_output_dir = _saved["output_dir"]

# ── sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 设置")
    st.text_input("工资表目录", key="merge_payroll_dir",
                   placeholder="例如: ./downloads/signed")
    st.text_input("输出目录(可选)", key="merge_output_dir",
                   placeholder="留空自动生成")
    st.caption("留空自动在工资表目录旁创建 merged_payrolls_{timestamp}")
    if st.button("💾 记住路径", use_container_width=True):
        _save_merge_settings(
            st.session_state.merge_payroll_dir.strip(),
            st.session_state.merge_output_dir.strip(),
        )
        st.success("路径已保存，下次启动自动加载", icon="✅")

    # ── Merge layout config ──
    with st.expander("📐 合并布局设置", expanded=False):
        _ml = _load_merge_layout_config()
        ml_changed = False
        
        _sig_h = st.number_input("签字行行高 (pt)", min_value=20, max_value=200, 
                                  value=_ml.get("pasted_sheet", {}).get("sig_row_height", 
                                        DEFAULT_MERGE_CONFIG["pasted_sheet"]["sig_row_height"]),
                                  key="ml_sig_h", help="合并后工资表签字行的行高")
        if _sig_h != _ml.get("pasted_sheet", {}).get("sig_row_height",
                          DEFAULT_MERGE_CONFIG["pasted_sheet"]["sig_row_height"]):
            ml_changed = True
        
        _id_w = st.number_input("身份证列宽", min_value=10, max_value=100,
                                 value=_ml.get("pasted_sheet", {}).get("id_card_min_width",
                                       DEFAULT_MERGE_CONFIG["pasted_sheet"]["id_card_min_width"]),
                                 key="ml_id_w", help="身份证号列最小列宽")
        if _id_w != _ml.get("pasted_sheet", {}).get("id_card_min_width",
                         DEFAULT_MERGE_CONFIG["pasted_sheet"]["id_card_min_width"]):
            ml_changed = True
        
        _sig_col = st.number_input("签字列最小列宽", min_value=5, max_value=100,
                                    value=_ml.get("pasted_sheet", {}).get("sig_col_min_width",
                                          DEFAULT_MERGE_CONFIG["pasted_sheet"]["sig_col_min_width"]),
                                    key="ml_sig_col", help="如总经理签字等列的列宽")
        if _sig_col != _ml.get("pasted_sheet", {}).get("sig_col_min_width",
                            DEFAULT_MERGE_CONFIG["pasted_sheet"]["sig_col_min_width"]):
            ml_changed = True
        
        _f_size = st.number_input("数据区字号 (pt)", min_value=6, max_value=20,
                                   value=_ml.get("font_size",
                                         DEFAULT_MERGE_CONFIG["font_size"]),
                                   key="ml_f_size", help="数据单元格字号")
        if _f_size != _ml.get("font_size", DEFAULT_MERGE_CONFIG["font_size"]):
            ml_changed = True
        
        _f_name = st.text_input("数据区字体",
                                 value=_ml.get("font_name",
                                       DEFAULT_MERGE_CONFIG["font_name"]),
                                 key="ml_f_name", help="数据单元格字体名称")
        if _f_name != _ml.get("font_name", DEFAULT_MERGE_CONFIG["font_name"]):
            ml_changed = True
        
        _lm = st.number_input("左边距 (cm)", min_value=0.0, max_value=5.0, step=0.1,
                               value=_ml.get("page", {}).get("left_margin_cm",
                                     DEFAULT_MERGE_CONFIG["page"]["left_margin_cm"]),
                               key="ml_lm", help="打印左边距")
        if _lm != _ml.get("page", {}).get("left_margin_cm",
                            DEFAULT_MERGE_CONFIG["page"]["left_margin_cm"]):
            ml_changed = True
        
        _rm = st.number_input("右边距 (cm)", min_value=0.0, max_value=5.0, step=0.1,
                               value=_ml.get("page", {}).get("right_margin_cm",
                                     DEFAULT_MERGE_CONFIG["page"]["right_margin_cm"]),
                               key="ml_rm", help="打印右边距")
        if _rm != _ml.get("page", {}).get("right_margin_cm",
                            DEFAULT_MERGE_CONFIG["page"]["right_margin_cm"]):
            ml_changed = True
        
        if ml_changed and st.button("💾 保存布局设置", key="ml_save", use_container_width=True):
            _save_merge_layout_config({
                "font_name": _f_name,
                "font_size": _f_size,
                "pasted_sheet": {
                    "sig_row_height": _sig_h,
                    "id_card_min_width": _id_w,
                    "sig_col_min_width": _sig_col,
                },
                "page": {
                    "left_margin_cm": _lm,
                    "right_margin_cm": _rm,
                },
            })
            st.success("布局设置已保存", icon="✅")

    st.divider()

    wps_ok = st.session_state.wps_available
    if wps_ok:
        st.info("WPS: ✅ 可用")
    else:
        wps_err = get_last_wps_error()
        if wps_err:
            st.error(f"⚠️ WPS 不可用\n\n`{wps_err}`")
        else:
            st.info("⚠️ WPS 不可用（Windows only）")

    st.divider()
    with st.expander("📊 汇总字段设置"):
        _sc_path = Path(__file__).parent.parent / "template" / "summary_config.json"
        _sc_default = {"fields": ["个人所得税", "个人工会会费", "工会经费", "实发合计", "实发工资", "扣工会会费"]}
        try:
            with open(_sc_path, "r", encoding="utf-8") as _f:
                _sc_cfg = json.load(_f)
        except (FileNotFoundError, json.JSONDecodeError):
            _sc_cfg = dict(_sc_default)
        _sc_fields = _sc_cfg.get("fields", list(_sc_default["fields"]))
        st.caption("编辑汇总字段（合并工资表的统计列）：")
        _new_fields = []
        for i, f in enumerate(_sc_fields):
            c1, c2 = st.columns([4, 1])
            _val = c1.text_input(f"字段 {i+1}", value=f, key=f"sc_f_{i}", label_visibility="collapsed")
            _new_fields.append(_val)
            if c2.button("🗑️", key=f"sc_del_{i}"):
                _new_fields.pop(i)
                _sc_cfg["fields"] = _new_fields
                with open(_sc_path, "w", encoding="utf-8") as _f:
                    json.dump(_sc_cfg, _f, ensure_ascii=False, indent=2)
                st.rerun()
        _new_name = st.text_input("新增字段名", key="sc_new_field", placeholder="输入字段名")
        if st.button("➕ 添加", key="sc_add", use_container_width=True) and _new_name.strip():
            if _new_name.strip() not in _new_fields:
                _new_fields.append(_new_name.strip())
                _sc_cfg["fields"] = _new_fields
                with open(_sc_path, "w", encoding="utf-8") as _f:
                    json.dump(_sc_cfg, _f, ensure_ascii=False, indent=2)
                st.rerun()
        if st.button("💾 保存配置", key="sc_save", use_container_width=True):
            _sc_cfg["fields"] = [f for f in _new_fields if f.strip()]
            with open(_sc_path, "w", encoding="utf-8") as _f:
                json.dump(_sc_cfg, _f, ensure_ascii=False, indent=2)
            st.success("已保存")

# ── main area ───────────────────────────────────────────────────────────────
st.title("🗂️ 工资表合并与打印")

# ── Merge Section ────────────────────────────────────────────────────────────
st.subheader("📂 合并工资表")
st.caption("将 signed_*.xlsx 工资表合并为一个汇总文件")

col_merge_btn, _ = st.columns([1, 3])
with col_merge_btn:
    merge_clicked = st.button("🔄 合并工资表", type="primary", use_container_width=True)

if merge_clicked:
    payroll_dir = st.session_state.get("merge_payroll_dir", "").strip()
    output_dir = st.session_state.get("merge_output_dir", "").strip()

    if not payroll_dir:
        st.error("请填写「工资表目录」")
    else:
        payroll_path = Path(payroll_dir)
        if not payroll_path.is_dir():
            st.error(f"目录不存在: {payroll_dir}")
        else:
            # 准备输出目录
            if not output_dir:
                from datetime import datetime
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = str(payroll_path.parent / f"merged_payrolls_{ts}")
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)

            # 进度显示
            progress_bar = st.progress(0, text="准备合并...")
            status_text = st.status("正在合并工资表...", expanded=True)

            def _progress_callback(current, total, msg):
                pct = current / total if total > 0 else 0
                progress_bar.progress(min(pct, 1.0), text=msg)
                status_text.write(msg)

            try:
                output_files, warnings, stats = merge_payrolls_simple(
                    str(payroll_path),
                    str(output_path),
                    progress_callback=_progress_callback,
                    merge_config=_load_merge_layout_config() or None,
                )

                status_text.update(label="合并完成", state="complete")
                progress_bar.empty()

                st.session_state.merge_results = {
                    "output_files": output_files,
                    "warnings": warnings,
                    "stats": stats,
                }

                count = stats.get("merged", len(output_files))
                st.success(f"✅ 合并完成：成功合并 {count} 个工资表，生成 {len(output_files)} 个文件")

                if output_files:
                    st.markdown("**输出文件：**")
                    for f in output_files:
                        st.markdown(f"- 📄 {f}")

                if warnings:
                    with st.expander(f"⚠️ 警告（{len(warnings)} 条）"):
                        for w in warnings:
                            st.warning(w)

            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                status_text.update(label="合并失败", state="error")
                progress_bar.empty()
                st.error(f"合并失败: {e}")
                with st.expander("错误详情"):
                    st.code(tb)

st.divider()

# ── Print Section (only when merge_results exists) ──────────────────────────
merge_results = st.session_state.merge_results

if merge_results:
    st.subheader("🖨️ 打印合并结果")

    output_files = merge_results.get("output_files", [])
    if not output_files:
        st.info("没有可打印的文件")
    else:
        st.caption(f"共 {len(output_files)} 个文件，勾选需要打印的文件")

        # 每个文件一个 checkbox
        selected_for_print = []
        for i, f in enumerate(output_files):
            checked = st.checkbox(f"📄 {f}", value=True, key=f"pf_{i}_{Path(f).name}")
            if checked:
                selected_for_print.append(f)

        st.divider()

        col_print_btn, col_info = st.columns([1, 2])
        with col_print_btn:
            print_clicked = st.button("🖨️ 打印所选文件", type="primary",
                                       use_container_width=True)
        with col_info:
            wps_ok = st.session_state.wps_available
            if wps_ok:
                st.caption("打印引擎: WPS / Excel COM")
            else:
                st.caption("打印引擎: LibreOffice")

        if print_clicked:
            if not selected_for_print:
                st.warning("请至少选择一个文件")
            else:
                print_status = st.status("正在打印...", expanded=True)
                print_progress = st.progress(0, text="准备打印...")

                def _print_callback(current, total, msg):
                    pct = current / total if total > 0 else 0
                    print_progress.progress(min(pct, 1.0), text=msg)
                    print_status.write(msg)

                try:
                    success, fail, fail_list = batch_print(
                        selected_for_print,
                        progress_callback=_print_callback,
                    )

                    print_progress.empty()

                    if success == 0 and fail == 0 and fail_list:
                        print_status.update(
                            label="打印失败：WPS 不可用",
                            state="error",
                        )
                        st.error("❌ WPS 不可用：无法启动 WPS COM 组件")
                        for msg in fail_list:
                            st.markdown(f"- {msg}")
                    elif success == 0 and fail == 0:
                        print_status.update(
                            label="打印完成：没有选择文件",
                            state="error",
                        )
                        st.warning("没有文件被打印")
                    elif fail == 0:
                        print_status.update(
                            label=f"打印完成：全部成功（{success} 个）",
                            state="complete",
                        )
                        st.success(f"✅ 全部 {success} 个文件打印成功")
                    else:
                        print_status.update(
                            label=f"打印完成：成功 {success} 个，失败 {fail} 个",
                            state="error" if success == 0 else "warning",
                        )
                        if success > 0:
                            st.success(f"✅ 成功打印 {success} 个文件")
                        if fail > 0:
                            st.error(f"❌ {fail} 个文件打印失败")
                            for f in fail_list:
                                st.markdown(f"- ❌ {f}")

                except Exception as e:
                    print_progress.empty()
                    print_status.update(label="打印失败", state="error")
                    st.error(f"打印过程异常: {e}")
