"""
Batch processor for approval instances.

Handles batch downloading of attachments from multiple approval instances.
"""

import json
import logging
import platform
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from .feishu_api import download_file, extract_attachments, get_instance_detail, parse_form

logger = logging.getLogger(__name__)

# Platform detection for print functions
HAS_WIN32COM = False
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    pass


def _print_with_com(file_path: Path, printer_name: Optional[str] = None) -> bool:
    """Print using WPS/Excel COM (Windows only)."""
    if platform.system() != "Windows":
        return False
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        app = None
        wb = None

        try:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            wb = app.Workbooks.Open(str(file_path.resolve()))
            if printer_name:
                wb.PrintOut(ActivePrinter=printer_name)
            else:
                wb.PrintOut()
            wb.Close(SaveChanges=False)
            app.Quit()
            return True
        except Exception as e:
            print(f"WPS/Excel打印失败: {e}")
            return False
        finally:
            try:
                if wb:
                    wb.Close(SaveChanges=False)
                if app:
                    app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
    except ImportError:
        return False


def _print_with_libreoffice(file_path: Path, printer_name: Optional[str] = None) -> bool:
    """Print using LibreOffice (cross-platform fallback)."""
    try:
        if file_path.suffix.lower() in (".xlsx", ".xls", ".docx", ".doc"):
            cmd = [
                "soffice",
                "--headless",
                "--print",
                str(file_path),
            ]
            if printer_name:
                cmd.extend(["--printer", printer_name])
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
            )
            return result.returncode == 0
        return False
    except FileNotFoundError:
        print("LibreOffice未安装")
        return False


def print_file(file_path: Path, printer_name: Optional[str] = None) -> bool:
    """Print file. Windows uses WPS/Excel COM, Linux uses LibreOffice."""
    if platform.system() == "Windows":
        success = _print_with_com(file_path, printer_name)
        if success:
            return True
        print("COM打印失败，尝试LibreOffice...")
        return _print_with_libreoffice(file_path, printer_name)
    else:
        return _print_with_libreoffice(file_path, printer_name)


_PAYROLL_CONFIG_PATH = Path(__file__).parent / "payroll_sheet_config.json"
_ROLE_MAPPING_PATH = Path(__file__).parent / "role_mapping.json"


def _load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _get_mandatory_roles() -> set:
    config = _load_json(_PAYROLL_CONFIG_PATH)
    mandatory = config.get("sheet_filter", {}).get("signatures", {}).get("mandatory", {})
    return {k for k in mandatory.keys() if k != "description"}


def _get_role_mapping(path: Optional[Path] = None) -> Dict[str, str]:
    mapping_path = path or _ROLE_MAPPING_PATH
    data = _load_json(mapping_path)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def is_ready_for_print(details: dict) -> bool:
    """
    Check if all mandatory roles in the approval have APPROVED status.

    Args:
        details: Instance detail dict containing "approver_list".

    Returns:
        True if every mandatory role (from payroll_sheet_config.json)
        has at least one APPROVED approver mapped to it.
    """
    mandatory_roles = _get_mandatory_roles()
    if not mandatory_roles:
        return False

    role_mapping = _get_role_mapping()
    approved_roles: set = set()

    for approver in details.get("approver_list", []):
        if approver.get("status") == "APPROVED":
            role = role_mapping.get(approver.get("approver_name"))
            if role:
                approved_roles.add(role)

    return mandatory_roles.issubset(approved_roles)


def is_approval_passed(details: dict) -> bool:
    """
    Check if the overall approval status is APPROVED.

    Args:
        details: Instance detail dict.

    Returns:
        True if details["status"] == "APPROVED".
    """
    return details.get("status") == "APPROVED"


def get_approvers_with_roles(details: dict, role_mapping_path: Optional[Path] = None) -> List[Dict[str, Any]]:
    """
    Extract approvers with their mapped roles from the approver_list.

    Args:
        details: Instance detail dict containing "approver_list".
        role_mapping_path: Optional path to a custom role_mapping.json file.

    Returns:
        List of dicts with keys "approver_name", "role", "status".
        If no role mapping is found for an approver_name, "role" is None.
    """
    role_mapping = _get_role_mapping(role_mapping_path)
    result: List[Dict[str, Any]] = []

    for approver in details.get("approver_list", []):
        name = approver.get("approver_name")
        role = role_mapping.get(name) if name else None
        result.append({
            "approver_name": name,
            "role": role,
            "status": approver.get("status"),
        })

    return result


_PAYROLL_CONFIG = None


def get_payroll_config() -> dict:
    """Load payroll sheet detection rules from config file."""
    global _PAYROLL_CONFIG
    if _PAYROLL_CONFIG is None:
        if _PAYROLL_CONFIG_PATH.exists():
            try:
                with open(_PAYROLL_CONFIG_PATH, "r", encoding="utf-8") as f:
                    _PAYROLL_CONFIG = json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                logger.warning(f"加载工资表配置失败: {e}，使用内置默认值")
        if _PAYROLL_CONFIG is None:
            _PAYROLL_CONFIG = {
                "sheet_filter": {
                    "row1_title": {"required_keyword": "工资发放表"},
                    "row2_org": {"required_keyword": "单位名称"},
                    "row3_headers": {"required": ["转账合计", "应发工资", "实发工资", "实发合计"]},
                    "signatures": {
                        "mandatory": {
                            "总经理签字": ["总经理签字"],
                            "部长签字": ["部长签字", "部长、分管副总签字", "分管副总签字"],
                            "财务审核": ["财务审核"],
                        },
                        "optional": {},
                    },
                }
            }
    return _PAYROLL_CONFIG


def is_payroll_sheet(ws, config: Optional[dict] = None) -> bool:
    """判断 worksheet 是否为应打印的工资发放表，规则全部从配置读取。"""
    cfg = config or get_payroll_config()
    sf = cfg["sheet_filter"]

    row1_text = ""
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        if cell.value:
            row1_text += str(cell.value).strip()
    req = sf["row1_title"]["required_keyword"]
    if req not in row1_text:
        return False

    mandatory = {
        k: v for k, v in sf["signatures"]["mandatory"].items()
        if k != "description"
    }
    found = set()
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell_val = str(ws.cell(row=row, column=col).value or "")
            for role, keywords in mandatory.items():
                if role not in found:
                    if any(kw in cell_val for kw in keywords):
                        found.add(role)
            if len(found) == len(mandatory):
                break
        if len(found) == len(mandatory):
            break
    if len(found) < len(mandatory):
        return False

    row2_text = ""
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        if cell.value:
            row2_text += str(cell.value).strip()
    normalized = row2_text.replace(" ", "").replace("\u3000", "")
    req_normalized = sf["row2_org"]["required_keyword"].replace(" ", "").replace("\u3000", "")
    if req_normalized not in normalized:
        return False

    required = set(sf["row3_headers"]["required"])
    row3_headers = set()
    for col in range(1, ws.max_column + 1):
        cell_val = str(ws.cell(row=3, column=col).value or "").strip()
        if cell_val:
            row3_headers.add(cell_val)
    if not required.issubset(row3_headers):
        return False

    return True


def _is_cell_in_merged_range(ws, row, col):
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
            return merged_range
    return None


def _split_merged_for_text(ws, row, col):
    from openpyxl.utils import get_column_letter

    merged = _is_cell_in_merged_range(ws, row, col)
    if not merged:
        cell = ws.cell(row=row, column=col)
        if cell.value and "部长、分管副总签字" in str(cell.value):
            cell.value = str(cell.value).replace("部长、分管副总签字", "部长签字")
        return col + 1

    cell = ws.cell(row=row, column=col)
    text = str(cell.value) if cell.value else ""

    if "部长、分管副总签字" in text:
        cell.value = text.replace("部长、分管副总签字", "部长签字")
        needed_cols = 3
    else:
        needed_cols = 2

    total_cols = merged.max_col - merged.min_col + 1
    if needed_cols >= total_cols:
        return merged.max_col + 1

    merged_str = str(merged)
    ws.unmerge_cells(merged_str)

    new_end = merged.min_col + needed_cols - 1
    ws.merge_cells(start_row=row, start_column=merged.min_col,
                   end_row=row, end_column=new_end)

    for c in range(new_end + 1, merged.max_col + 1):
        ws.cell(row=row, column=c).value = None

    return new_end + 1


def find_all_signature_positions(ws, config: Optional[dict] = None) -> Dict[str, Tuple[int, int]]:
    positions = {}
    cfg = config or get_payroll_config()
    sf = cfg["sheet_filter"]["signatures"]

    keyword_groups = []
    for role, keywords in sf.get("mandatory", {}).items():
        keyword_groups.append((keywords, role))
    for role, keywords in sf.get("optional", {}).items():
        keyword_groups.append((keywords, role))

    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                value = str(cell.value).strip()
                for texts, role_key in keyword_groups:
                    if role_key not in positions:
                        for text in texts:
                            if text in value:
                                positions[role_key] = (row, col)
                                break
    return positions


def _extract_first_row_title(ws) -> Optional[str]:
    """Find the first non-empty cell in row 1 as the table title."""
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        if cell.value:
            return str(cell.value).strip()
    return None


def _build_output_path(excel_path: Path, output_path: Path, ws) -> Path:
    """If the original file is an unrenamed export (tddd_dialog*), append row-1 title."""
    original_name = excel_path.name
    if original_name.lower().startswith("tddd_dialog"):
        title = _extract_first_row_title(ws)
        if title:
            safe_title = sanitize_dir_name(title)
            new_name = f"signed_{Path(original_name).stem}-{safe_title}.xlsx"
            return output_path.parent / new_name
    return output_path


def _find_total_row(ws) -> int:
    keywords = ["合计", "总计", "合计金额", "合计费用", "合计支付"]
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                val = str(cell.value).strip()
                for kw in keywords:
                    if kw in val:
                        return row
    return 0


def _apply_border_styles(ws, signature_positions):
    from openpyxl.styles import Border, Side
    try:
        thick = Side(style='medium', color='000000')
        none = Side(style=None)

        total_row = _find_total_row(ws)
        sig_rows = sorted({r for (r, _) in signature_positions.values()})
        if not sig_rows:
            logger.info("[BORDER] 无签名位置，跳过边框设置")
            return
        sig_start = min(sig_rows)
        sig_end = max(sig_rows)
        last_col = ws.max_column

        border_start = 3
        border_end = sig_end + 2

        if border_start <= border_end:
            for r in range(border_start, border_end + 1):
                for c in range(1, last_col + 1):
                    is_top = r == border_start
                    is_bottom = r == border_end
                    is_left = c == 1
                    is_right = c == last_col
                    if not (is_top or is_bottom or is_left or is_right):
                        continue
                    cell = ws.cell(row=r, column=c)
                    cell.border = Border(
                        top=thick if is_top else cell.border.top,
                        bottom=thick if is_bottom else cell.border.bottom,
                        left=thick if is_left else cell.border.left,
                        right=thick if is_right else cell.border.right,
                    )

        logger.info(f"[BORDER] 外框从第{border_start}行到第{border_end}行，内部单元格边框保留")
    except Exception as e:
        logger.warning(f"[BORDER] 边框设置出错: {e}")


def _hide_columns(ws):
    headers_to_hide = {"部门", "岗位", "职工号"}
    header_row = 3
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col)
        col_letter = cell.column_letter
        hidden = ws.column_dimensions[col_letter].hidden
        if cell.value and str(cell.value).strip() in headers_to_hide:
            hidden = True
        ws.column_dimensions[col_letter].hidden = hidden


def adjust_excel_for_print(ws, signature_positions=None) -> None:
    """
    调整 Excel 打印设置：横向打印，A4 纸，左边距 2cm，其他边距 1cm，
    所有列缩放到 1 页宽，水平居中。
    在嵌入签名图片前调用。
    """
    try:
        ws.page_setup.paperSize = 9          # A4
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True  # 正确路径，必须先启用
        ws.page_setup.fitToWidth = 1         # 缩放到 1 页宽
        ws.page_setup.fitToHeight = 0        # 高度自动分页
        ws.page_margins.left = 0.8           # 2cm
        ws.page_margins.right = 0.4          # 1cm
        ws.page_margins.top = 0.4            # 1cm
        ws.page_margins.bottom = 0.4         # 1cm
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = False
        ws.print_options.gridLines = False
        logger.info("[PRINT] 已调整: 横向A4, 左2cm其余1cm, 1页宽, fitToPage=True, 网格线关闭")

        _hide_columns(ws)

        if signature_positions:
            _apply_border_styles(ws, signature_positions)
    except Exception as e:
        logger.warning(f"[PRINT] 调整打印设置时出错: {e}")


def _insert_signature_to_excel_openpyxl(
    excel_path: Path,
    approvers: List[Dict],
    signatures_dir: Path,
    output_path: Path,
) -> Tuple[bool, List[str], Path]:
    inserted_roles = []
    try:
        logger.info(f"[SIGN] Loading workbook: {excel_path.name}")
        wb = load_workbook(str(excel_path))

        cfg = get_payroll_config()
        payroll_ws = None
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if is_payroll_sheet(ws, cfg):
                payroll_ws = ws
                logger.info(f"[SIGN] Found payroll sheet: {sheet_name}")
                break

        if payroll_ws is None:
            logger.warning(f"[SIGN] No payroll sheet found in {excel_path.name}")
            return False, [], output_path

        positions = find_all_signature_positions(payroll_ws, cfg)
        adjust_excel_for_print(payroll_ws, positions)
        logger.info(f"[SIGN] Found positions: {positions}")
        if not positions:
            logger.warning(f"[SIGN] No signature positions found in {excel_path.name}")
            return False, [], output_path

        logger.info(f"[SIGN] Approvers: {[a['role'] for a in approvers]}")
        for approver in approvers:
            role = approver.get("role")
            approver_name = approver.get("approver_name")
            if not role or not approver_name:
                logger.info(f"[SIGN] Skipping approver: missing role or approver_name")
                continue

            if role not in positions:
                logger.info(f"[SIGN] Role '{role}' not in positions")
                continue

            sig_path = get_signature_path(approver_name, signatures_dir)
            if not sig_path:
                logger.warning(f"[SIGN] Signature image not found for {approver_name}")
                continue

            row, col = positions[role]
            target_col = _split_merged_for_text(payroll_ws, row, col)

            img = XLImage(str(sig_path))
            img.width = 120
            img.height = 60
            cell_addr = f"{payroll_ws.cell(row=row, column=target_col).coordinate}"
            payroll_ws.add_image(img, cell_addr)
            inserted_roles.append(role)
            logger.info(f"[SIGN] Inserted signature for {role} at {cell_addr}")

        actual_output = _build_output_path(excel_path, output_path, payroll_ws)

        sheets_to_remove = [sn for sn in wb.sheetnames if sn != payroll_ws.title]
        for sn in sheets_to_remove:
            del wb[sn]
            logger.info(f"[SIGN] Removed non-payroll sheet: {sn}")

        wb.save(str(actual_output))
        logger.info(f"[SIGN] Saved to {actual_output.name}, inserted: {inserted_roles}")
        return len(inserted_roles) > 0, inserted_roles, actual_output
    except Exception as e:
        logger.error(f"[SIGN] Insertion failed: {e}")
        return False, [], output_path


def sanitize_dir_name(name: str) -> str:
    """Sanitize string for use as directory name."""
    invalid_chars = '\\/:*?"<>|'
    for char in invalid_chars:
        name = name.replace(char, "_")
    return name.strip()


def get_signature_path(approver_name: str, signatures_dir: Path) -> Optional[Path]:
    """Look up signature PNG file for an approver by name."""
    if not approver_name:
        return None
    sig_path = signatures_dir / f"{approver_name}.png"
    if sig_path.exists():
        return sig_path
    for f in signatures_dir.glob("*.png"):
        if approver_name in f.name:
            return f
    return None


def process_single_approval(
    instance_code: str,
    token: str,
    config: dict,
) -> dict:
    """Process a single approval: download, sign, print.

    Args:
        instance_code: Feishu approval instance code.
        token: Feishu tenant access token.
        config: Dict with keys:
            - save_dir: download directory (default: "./downloads")
            - signatures_dir: signature images directory (default: "./signatures")
            - role_mapping_path: optional path to role mapping json

    Returns:
        Result dict with keys: instance_code, success, message, downloaded,
        signed, signed_files, skipped, title.
    """
    result = {
        "instance_code": instance_code,
        "success": False,
        "message": "",
        "downloaded": [],
        "signed": [],
        "signed_files": [],
        "skipped": False,
        "title": "",
    }

    try:
        logger.info(f"[BATCH] Getting details for {instance_code}...")
        detail = get_instance_detail(token, instance_code)
        logger.info(f"[BATCH] Got details from API")
    except Exception as e:
        logger.error(f"[BATCH] Failed to get details: {e}")
        result["message"] = f"获取详情失败: {e}"
        return result

    result["title"] = detail.get("title", instance_code[:20])

    if not is_approval_passed(detail):
        result["skipped"] = True
        result["message"] = "审批未通过"
        return result

    role_mapping_path = config.get("role_mapping_path")
    if role_mapping_path:
        role_mapping_path = Path(role_mapping_path)
    approvers = get_approvers_with_roles(detail, role_mapping_path)
    approvers = [a for a in approvers if a.get("status") == "APPROVED" and a.get("role")]
    if not approvers:
        result["message"] = "未找到可签名的审批角色"
        return result

    form_widgets = parse_form(detail)
    attachments = extract_attachments(form_widgets)
    if not attachments:
        result["message"] = "无附件"
        return result

    save_dir = Path(config.get("save_dir", "./downloads"))
    instance_dir = save_dir / sanitize_dir_name(instance_code)
    instance_dir.mkdir(parents=True, exist_ok=True)

    signatures_dir = Path(config.get("signatures_dir", "./signatures"))

    for att in attachments:
        field_name = att.get("field_name", "附件")
        values = att.get("value", [])
        if not values:
            continue

        if "汇总表" in field_name:
            logger.info(f"[BATCH] Skipping summary table: {field_name}")
            continue

        file_token = values[0]

        try:
            logger.info(f"[BATCH] Processing attachment: {field_name}")
            downloaded_path = download_file(token, file_token, str(instance_dir))
            file_path = Path(downloaded_path)
            result["downloaded"].append(file_path.name)

            if file_path.suffix.lower() in (".xlsx", ".xls"):
                logger.info(f"[BATCH] Inserting signatures into {file_path.name}...")
                signed_name = f"signed_{file_path.stem}.xlsx"
                signed_path = instance_dir / signed_name
                success, inserted, actual_signed_path = _insert_signature_to_excel_openpyxl(
                    file_path, approvers, signatures_dir, signed_path
                )
                if success:
                    logger.info(f"[BATCH] Signature insertion success: {inserted}")
                    result["signed"].extend(inserted)
                    result["signed_files"].append(str(actual_signed_path))
                else:
                    logger.warning(f"[BATCH] Signature insertion failed for {file_path.name}")
            else:
                logger.info(f"[BATCH] Non-Excel file, skipping signature: {file_path.name}")

        except Exception as e:
            logger.error(f"[BATCH] Error processing {field_name}: {e}")
            continue

    result["success"] = True
    result["message"] = (
        f"下载 {len(result['downloaded'])} 个, "
        f"签名 {len(result['signed'])} 处"
    )
    return result
