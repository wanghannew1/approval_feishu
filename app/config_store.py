"""统一 JSON 配置存储。

设计目标:
- 所有用户可编辑的 .json 配置文件都不进版本库 (见 .gitignore)；
- 默认值集中在本模块以 Python 字面量形式定义；
- 首次读取时若磁盘文件不存在，自动写入默认值并返回，保证生产环境
  ``git clone`` 后空目录即可运行、UI 立即出现可编辑的默认配置文件。

对外主要入口:
- :func:`load_or_create`  —— 读，缺失则落盘默认值并返回。
- :func:`save`            —— 写 (兼容旧版 ``_save_*`` 调用点)。
- :func:`ensure_all`      —— 一次性把所有已知配置文件的默认值落盘。
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Dict

logger = logging.getLogger(__name__)

# ────────────────────────────── 默认值 ──────────────────────────────────────
# 注意: 任何默认值变更都要同步这里的 Python 字面量。

DEFAULT_DOWNLOAD_PATH = r"C:\Users\BY\Desktop\当天工资"

DEFAULT_DEFINITIONS: Dict[str, str] = {
    "1CF34ABB-781C-40B0-9A4F-3CC416612423": "项目人员工资发放审批单（系统工资单）",
}

DEFAULT_SETTINGS: Dict[str, object] = {
    "download_path": DEFAULT_DOWNLOAD_PATH,
    "merge_payroll_dir": DEFAULT_DOWNLOAD_PATH,
    "merge_output_dir": "",
}

# 用户 id -> 姓名: 新用户首次需要手动维护 (或导入 Excel)，默认空。
DEFAULT_USER_MAPPING: Dict[str, str] = {}

# 飞书审批节点名 -> Excel 签名位置关键词。
DEFAULT_ROLE_MAPPING: Dict[str, str] = {
    "业务审核": "业务审核",
    "部长签字": "分管领导审核",
    "分管领导审核": "分管领导审核",
    "财务审核": "财务审核",
    "总经理签字": "总经理签字",
    "部门负责人": "部长签字",
    "总经理": "总经理签字",
    "财务": "财务审核",
    "五险一金": "业务审核",
}

# 审批记录显示顺序，序号越小越靠前。
DEFAULT_WORKFLOW_ORDER: Dict[str, int] = {
    "提交": 1,
    "业务审核": 2,
    "分管领导审核": 3,
    "财务审核": 4,
    "总经理签字": 5,
    "出纳办理": 6,
    "结束": 99,
}

# 工资发放表识别 / 文本归一化 / 强删列 / 空列清理 / 打印布局。
DEFAULT_PAYROLL_CONFIG: Dict[str, object] = {
    "sheet_filter": {
        "title_keyword": {"required": "工资"},
        "exclude_keywords": {"keywords": ["汇总数据"]},
        "required_content": {"required": ["应发工资", "实发工资"]},
        "signatures": {
            "mandatory": {"总经理签字": ["总经理签字"]},
            "optional": {
                "分管领导审核": ["分管领导审核"],
                "财务审核": ["财务审核"],
                "业务审核": ["业务审核"],
            },
        },
    },
    "text_normalization": {
        "rules": [
            {"source": "部长、分管副总签字", "target": "分管领导审核"},
            {"source": "部长签字", "target": "分管领导审核"},
            {"source": "分管领导签字", "target": "分管领导审核"},
        ]
    },
    "force_delete_columns": {"columns": ["岗位"]},
    "remove_empty_columns": {"enabled": True},
    "layout": {"id_card_min_width": 20},
}

# 合并工资表统计字段配置 (template/summary_config.json)。
DEFAULT_SUMMARY_CONFIG: Dict[str, object] = {
    "fields": [
        "个人所得税",
        "个人工会会费",
        "工会经费",
        "实发合计",
        "实发工资",
        "扣工会会费",
    ]
}

# template/validation_config.json 目前没有被代码读取，但保留默认值以便
# 将来引入校验逻辑或人工对照。
DEFAULT_VALIDATION_CONFIG: Dict[str, object] = {
    "tolerance": 0.005,
    "deduction": {"prefix": "扣款明细/", "exclude_keywords": ["扣款合计", "大病险合计"]},
}


# ────────────────────────────── 路径 ──────────────────────────────────────
_APP_DIR = Path(__file__).parent

PATH_DEFINITIONS = _APP_DIR / "approval_definitions.json"
PATH_SETTINGS = _APP_DIR / "settings.json"
PATH_USER_MAPPING = _APP_DIR / "user_mapping.json"
PATH_PAYROLL_CONFIG = _APP_DIR / "payroll_sheet_config.json"
PATH_ROLE_MAPPING = _APP_DIR / "role_mapping.json"
PATH_WORKFLOW_ORDER = _APP_DIR / "workflow_order.json"
PATH_SUMMARY_CONFIG = _APP_DIR / "template" / "summary_config.json"
PATH_VALIDATION_CONFIG = _APP_DIR / "template" / "validation_config.json"

_ALL_DEFAULTS: Dict[Path, Dict[str, object]] = {
    PATH_DEFINITIONS: DEFAULT_DEFINITIONS,
    PATH_SETTINGS: DEFAULT_SETTINGS,
    PATH_USER_MAPPING: DEFAULT_USER_MAPPING,
    PATH_PAYROLL_CONFIG: DEFAULT_PAYROLL_CONFIG,
    PATH_ROLE_MAPPING: DEFAULT_ROLE_MAPPING,
    PATH_WORKFLOW_ORDER: DEFAULT_WORKFLOW_ORDER,
    PATH_SUMMARY_CONFIG: DEFAULT_SUMMARY_CONFIG,
    PATH_VALIDATION_CONFIG: DEFAULT_VALIDATION_CONFIG,
}


# ────────────────────────────── API ───────────────────────────────────────
def load_or_create(path: Path, default: Dict[str, object]) -> Dict[str, object]:
    """读取 JSON 配置；文件缺失则写入 ``default`` 并返回其深拷贝。

    解析失败 (json.JSONDecodeError) 时记录警告并回退到 ``default``，
    但不覆盖磁盘上损坏的文件 —— 避免意外清掉用户手改了一半的内容。
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        logger.info("配置文件缺失，写入默认值: %s", path)
        save(path, default)
        return json.loads(json.dumps(default, ensure_ascii=False))  # deep copy
    except json.JSONDecodeError as e:
        logger.warning("配置文件解析失败 (%s): %s —— 使用默认值", path, e)
        return json.loads(json.dumps(default, ensure_ascii=False))


def save(path: Path, data: Dict[str, object]) -> None:
    """以 UTF-8、缩进 2、不转义中文的方式写回 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def ensure_all() -> None:
    """把所有已知配置文件的默认值落盘 (仅当文件不存在时)。"""
    for path, default in _ALL_DEFAULTS.items():
        if not path.exists():
            save(path, default)
            logger.info("写入默认配置: %s", path)