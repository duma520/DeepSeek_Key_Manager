#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 DeepSeek 密钥余额管理器  (DeepSeek Key Manager)
================================================================================
 功能概述
 --------
 • 批量管理**无上限**数量的 DeepSeek API Key（用户提供多少个就列出多少个的信息）
 • 多线程并发调用官方接口 https://api.deepseek.com/user/balance
 • 表格完整展示接口原始字段：
       Deepseek API KEY / is_available(账户是否可用) / balance_infos(余额明细)
       / total_balance(总可用余额) / granted_balance(赠金余额)
       / topped_up_balance(充值余额) / currency(货币单位)
   并额外补充：别名备注、查询状态(带进度条)、最后检查时间、耗时、错误说明
 • 每个密钥一行进度条 + 底部总体进度条，实时显示当前进展
 • 所有设置实时自动保存（防抖）+ 关闭时强制保存，下次启动自动恢复
 • 全部设置与密钥存放在程序目录：.env（密钥/数据库密码）与 *_settings.json
 • 支持 导入(.env / txt / json / 剪贴板)、导出(CSV / Excel / JSON / Markdown)
 • 支持 代理、超时、重试、并发数、SSL 校验、自动刷新、明暗主题、显示列设置
 • 支持 Nuitka 编译（唯一编译源即本文件）

 安全说明
 --------
 • API Key 与数据库密码一律存放于 .env，**不硬编码进代码**
 • 界面上默认对 Key 做掩码显示；导出时可选择是否包含完整 Key

 依赖：PySide6 / requests / (可选) openpyxl
================================================================================
"""

from __future__ import annotations

# ============================================================================
#  区域 0：元信息常量
# ============================================================================

__version__ = "1.1.1"
__author__ = "GitHub Copilot"
__product__ = "DeepSeek 密钥余额管理器"
__product_en__ = "DeepSeek Key Manager"

DSKM_VERSION = __version__
DSKM_APP_NAME = __product__
DSKM_ORG_NAME = "DeepSeekKeyManager"
DSKM_APP_ID = "DeepSeekKeyManager"

DSKM_DEFAULT_API_BASE = "https://api.deepseek.com"
DSKM_BALANCE_PATH = "/user/balance"
DSKM_TOPUP_URL = "https://platform.deepseek.com/top_up"
DSKM_USAGE_URL = "https://platform.deepseek.com/usage"
DSKM_CONFIG_URL = "https://platform.deepseek.com/api_keys"

DSKM_USER_AGENT = f"DeepSeekKeyManager/{__version__} (Windows GUI)"

# 接口字段（与官方文档保持一致，界面上中英对照）
DSKM_FIELD_IS_AVAILABLE = "is_available"
DSKM_FIELD_BALANCE_INFOS = "balance_infos"
DSKM_FIELD_TOTAL_BALANCE = "total_balance"
DSKM_FIELD_GRANTED_BALANCE = "granted_balance"
DSKM_FIELD_TOPPED_UP_BALANCE = "topped_up_balance"
DSKM_FIELD_CURRENCY = "currency"

# ---- API Key 在列表/导出/剪贴板里的显示方式（v1.1.0 新增）----
KEY_DISPLAY_FULL = "full"              # 完整显示（不掩码）
KEY_DISPLAY_HEAD_TAIL = "head_tail"    # 前 N 位 + 掩码 + 后 M 位（默认）
KEY_DISPLAY_HEAD_ONLY = "head_only"    # 只显示前 N 位，其余全部掩码
KEY_DISPLAY_TAIL_ONLY = "tail_only"    # 只显示后 M 位，前面全部掩码
KEY_DISPLAY_ALL_MASK = "all_mask"      # 全部用掩码符号

KEY_DISPLAY_MODE_OPTIONS: List[Tuple[str, str]] = [
    (KEY_DISPLAY_HEAD_TAIL, "前 N 位 + 后 M 位（推荐）"),
    (KEY_DISPLAY_HEAD_ONLY, "只显示前 N 位"),
    (KEY_DISPLAY_TAIL_ONLY, "只显示后 M 位"),
    (KEY_DISPLAY_ALL_MASK, "全部用掩码符号（最安全）"),
    (KEY_DISPLAY_FULL, "完整显示（不隐藏）"),
]

KEY_DISPLAY_MODE_LABELS: Dict[str, str] = {mode: label for mode, label in KEY_DISPLAY_MODE_OPTIONS}

KEY_MASK_CHAR_OPTIONS: List[Tuple[str, str]] = [
    ("*", "* 星号（默认）"),
    ("•", "• 圆点"),
    ("●", "● 实心圆"),
    ("×", "× 叉号"),
    ("#", "# 井号"),
    ("·", "· 中间点"),
]

KEY_DISPLAY_SAMPLE = "sk-***已脱敏***"

# 设置键名常量（避免手写字符串写错）
SET_KEY_WINDOW_GEOMETRY = "window_geometry"
SET_KEY_WINDOW_MAXIMIZED = "window_maximized"
SET_KEY_WINDOW_ON_TOP = "window_always_on_top"
SET_KEY_SPLITTER_SIZES = "splitter_sizes"
SET_KEY_LOG_VISIBLE = "log_visible"
SET_KEY_LOG_HEIGHT = "log_height"
SET_KEY_LOG_MAX_LINES = "log_max_lines"
SET_KEY_LOG_AUTOSCROLL = "log_autoscroll"
SET_KEY_COLUMN_WIDTHS = "column_widths"
SET_KEY_HIDDEN_COLUMNS = "hidden_columns"
SET_KEY_COLUMN_ORDER = "column_order"
SET_KEY_SORT_COLUMN = "sort_column"
SET_KEY_SORT_ORDER = "sort_order"
SET_KEY_KEY_DISPLAY_MODE = "key_display_mode"
SET_KEY_KEY_DISPLAY_HEAD = "key_display_head"
SET_KEY_KEY_DISPLAY_TAIL = "key_display_tail"
SET_KEY_KEY_MASK_CHAR = "key_mask_char"
SET_KEY_KEY_MASK_LENGTH = "key_mask_length"
SET_KEY_LAST_MASKED_MODE = "last_masked_mode"
# 已废弃（v1.1.0 起改用 key_display_mode；启动时自动迁移，见 DSKMSettingsStore.load）
SET_KEY_MASK_API_KEY = "mask_api_key"
SET_KEY_MASK_FULL = "mask_full_hide"
SET_KEY_API_BASE = "api_base"
SET_KEY_TIMEOUT = "timeout"
SET_KEY_MAX_RETRY = "max_retry"
SET_KEY_CONCURRENCY = "concurrency"
SET_KEY_VERIFY_SSL = "verify_ssl"
SET_KEY_PROXY_ENABLED = "proxy_enabled"
SET_KEY_PROXY_URL = "proxy_url"
SET_KEY_AUTO_REFRESH = "auto_refresh"
SET_KEY_AUTO_REFRESH_INTERVAL = "auto_refresh_interval"
SET_KEY_CHECK_ON_STARTUP = "check_on_startup"
SET_KEY_ONLY_ENABLED_CHECK = "only_enabled_check"
SET_KEY_SKIP_RECENT_CHECK = "skip_recent_check"
SET_KEY_SKIP_RECENT_SECONDS = "skip_recent_seconds"
SET_KEY_THEME = "theme"
SET_KEY_FONT_SIZE = "font_size"
SET_KEY_ROW_HEIGHT = "row_height"
SET_KEY_CURRENCY_PRECISION = "currency_precision"
SET_KEY_EXPORT_DIR = "export_dir"
SET_KEY_IMPORT_DIR = "import_dir"
SET_KEY_LAST_EXPORT_FORMAT = "last_export_format"
SET_KEY_CONFIRM_ON_DELETE = "confirm_on_delete"
SET_KEY_CONFIRM_ON_CLEAR = "confirm_on_clear"
SET_KEY_KEEP_HISTORY = "keep_history"
SET_KEY_HISTORY_LIMIT = "history_limit"
SET_KEY_STATISTICS_BAR = "statistics_bar_visible"
SET_KEY_SEARCH_TEXT = "search_text"
SET_KEY_STATUS_FILTER = "status_filter"
SET_KEY_CURRENCY_FILTER = "currency_filter"
SET_KEY_EXPORT_INCLUDE_KEY = "export_include_key"
SET_KEY_NOTIFY_ON_FINISH = "notify_on_finish"
SET_KEY_AUTO_DEDUPE = "auto_dedupe"

DEFAULT_SETTINGS: dict = {
    SET_KEY_WINDOW_GEOMETRY: [560, 120, 1420, 860],
    SET_KEY_WINDOW_MAXIMIZED: False,
    SET_KEY_WINDOW_ON_TOP: False,
    SET_KEY_SPLITTER_SIZES: [],
    SET_KEY_LOG_VISIBLE: True,
    SET_KEY_LOG_HEIGHT: 150,
    SET_KEY_LOG_MAX_LINES: 2000,
    SET_KEY_LOG_AUTOSCROLL: True,
    SET_KEY_COLUMN_WIDTHS: {},
    SET_KEY_HIDDEN_COLUMNS: ["elapsed_ms"],
    SET_KEY_COLUMN_ORDER: [],
    SET_KEY_SORT_COLUMN: "",
    SET_KEY_SORT_ORDER: "",
    SET_KEY_KEY_DISPLAY_MODE: KEY_DISPLAY_HEAD_TAIL,
    SET_KEY_KEY_DISPLAY_HEAD: 6,
    SET_KEY_KEY_DISPLAY_TAIL: 4,
    SET_KEY_KEY_MASK_CHAR: "*",
    SET_KEY_KEY_MASK_LENGTH: 8,
    SET_KEY_LAST_MASKED_MODE: KEY_DISPLAY_HEAD_TAIL,
    SET_KEY_API_BASE: DSKM_DEFAULT_API_BASE,
    SET_KEY_TIMEOUT: 20,
    SET_KEY_MAX_RETRY: 2,
    SET_KEY_CONCURRENCY: 5,
    SET_KEY_VERIFY_SSL: True,
    SET_KEY_PROXY_ENABLED: False,
    SET_KEY_PROXY_URL: "",
    SET_KEY_AUTO_REFRESH: False,
    SET_KEY_AUTO_REFRESH_INTERVAL: 30,
    SET_KEY_CHECK_ON_STARTUP: False,
    SET_KEY_ONLY_ENABLED_CHECK: True,
    SET_KEY_SKIP_RECENT_CHECK: False,
    SET_KEY_SKIP_RECENT_SECONDS: 300,
    SET_KEY_THEME: "light",
    SET_KEY_FONT_SIZE: 9,
    SET_KEY_ROW_HEIGHT: 30,
    SET_KEY_CURRENCY_PRECISION: 2,
    SET_KEY_EXPORT_DIR: "",
    SET_KEY_IMPORT_DIR: "",
    SET_KEY_LAST_EXPORT_FORMAT: "csv",
    SET_KEY_CONFIRM_ON_DELETE: True,
    SET_KEY_CONFIRM_ON_CLEAR: True,
    SET_KEY_KEEP_HISTORY: True,
    SET_KEY_HISTORY_LIMIT: 200,
    SET_KEY_STATISTICS_BAR: True,
    SET_KEY_SEARCH_TEXT: "",
    SET_KEY_STATUS_FILTER: "all",
    SET_KEY_CURRENCY_FILTER: "all",
    SET_KEY_EXPORT_INCLUDE_KEY: False,
    SET_KEY_NOTIFY_ON_FINISH: True,
    SET_KEY_AUTO_DEDUPE: True,
}

# 状态常量
STATUS_PENDING = "pending"
STATUS_RUNNING = "running"
STATUS_OK = "ok"
STATUS_FAIL = "fail"
STATUS_DISABLED = "disabled"

STATUS_TEXT = {
    STATUS_PENDING: "待查询",
    STATUS_RUNNING: "查询中…",
    STATUS_OK: "查询成功",
    STATUS_FAIL: "查询失败",
    STATUS_DISABLED: "已停用",
}

# HTTP 错误提示（中文）
DSKM_HTTP_HINTS = {
    400: "请求参数错误（400）",
    401: "认证失败：API Key 无效、被禁用或已被删除（401）",
    402: "余额不足（402）",
    403: "权限不足，无法访问该接口（403）",
    404: "接口地址不存在，请检查 API 基地址（404）",
    405: "请求方法不被允许（405）",
    422: "请求格式错误（422）",
    429: "请求过于频繁，请降低并发数或稍后重试（429）",
    500: "DeepSeek 服务端内部错误（500）",
    502: "网关错误（502）",
    503: "服务暂时不可用（503）",
    504: "网关超时（504）",
}

DSKM_LOG_LEVELS = {
    "debug": "调试",
    "info": "信息",
    "success": "成功",
    "warn": "警告",
    "error": "错误",
}

# ============================================================================
#  区域 1：导入
# ============================================================================

import csv
import datetime as _dt
import hashlib
import json
import os
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from PySide6.QtCore import (
    QAbstractTableModel,
    QEvent,
    QModelIndex,
    QObject,
    QPoint,
    QRect,
    QRunnable,
    QSize,
    QSortFilterProxyModel,
    QThreadPool,
    QTimer,
    Qt,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QAction,
    QActionGroup,
    QBrush,
    QColor,
    QFont,
    QFontDatabase,
    QGuiApplication,
    QIcon,
    QKeySequence,
    QPainter,
    QPalette,
    QPen,
    QPixmap,
    QShortcut,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFontDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QDoubleSpinBox,
    QSplitter,
    QStackedWidget,
    QStyle,
    QStyleOptionViewItem,
    QStyledItemDelegate,
    QTabWidget,
    QScrollArea,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTextEdit,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

try:  # 可选依赖：网络请求
    import requests
    from requests.adapters import HTTPAdapter

    DSKM_HAS_REQUESTS = True
except Exception:  # pragma: no cover - 环境缺失时给出明确提示
    requests = None  # type: ignore[assignment]
    HTTPAdapter = None  # type: ignore[assignment]
    DSKM_HAS_REQUESTS = False

def dskm_enum_value(*getters: Callable[[], Any], default: Any = None) -> Any:
    """按顺序尝试取枚举值，返回第一个可用的（兼容 Qt5 扁平写法与 Qt6 枚举类写法）。

    ★ 踩过的坑（v1.1.1 修）：以前把这些常量全塞进一个 `try: ... except: ...`，
    只要其中**任意一个**名字在当前 PySide6 里不存在，`except` 就会把前面所有
    已经取到值的常量**一起覆盖成 fallback**。实测 PySide6 6.11 已经移除了
    `Qt.TextColorRole`（Qt6 里必须写成 `Qt.ItemDataRole.TextColorRole`），
    于是 `DSKM_STYLE_NO_PEN` 被连带改成 None，委托每次重绘都报
    `QPainter.setPen called with wrong argument values: (None,)`，
    控制台还会刷 `QPainter::end: Painter ended with 1 saved states`。
    ⇒ **每个常量必须独立求解，绝不共用一个大 try。**
    """
    for getter in getters:
        try:
            value = getter()
        except Exception:
            continue
        if value is not None:
            return value
    return default


# 只保留真正用到的枚举常量（现代写法优先，扁平写法兜底）
DSKM_PEN_NO_PEN = dskm_enum_value(lambda: Qt.PenStyle.NoPen, lambda: Qt.NoPen, default=1)
DSKM_BRUSH_NO_BRUSH = dskm_enum_value(lambda: Qt.BrushStyle.NoBrush, lambda: Qt.NoBrush, default=0)
DSKM_PEN_DOT_LINE = dskm_enum_value(lambda: Qt.PenStyle.DotLine, lambda: Qt.DotLine, default=3)
DSKM_STATE_SELECTED = dskm_enum_value(
    lambda: QStyle.StateFlag.State_Selected, lambda: QStyle.State_Selected, default=0
)
DSKM_STATE_MOUSE_OVER = dskm_enum_value(
    lambda: QStyle.StateFlag.State_MouseOver, lambda: QStyle.State_MouseOver, default=0
)
DSKM_CURSOR_WAIT = dskm_enum_value(
    lambda: Qt.CursorShape.WaitCursor, lambda: Qt.WaitCursor, default=3
)


try:  # 可选依赖：Excel 导出
    import openpyxl

    DSKM_HAS_OPENPYXL = True
except Exception:  # pragma: no cover
    openpyxl = None  # type: ignore[assignment]
    DSKM_HAS_OPENPYXL = False


# ============================================================================
#  区域 2：路径与通用工具
# ============================================================================


def dskm_is_compiled() -> bool:
    """判断当前是否运行在 Nuitka / PyInstaller 打包环境中。"""
    if "__compiled__" in globals():
        return True
    if getattr(sys, "frozen", False):
        return True
    return False


def dskm_app_dir() -> Path:
    """程序所在目录（开发态为脚本目录，打包态为 exe 目录）。"""
    if dskm_is_compiled():
        try:
            return Path(sys.executable).resolve().parent
        except Exception:
            pass
    try:
        return Path(os.path.abspath(__file__)).resolve().parent
    except NameError:  # pragma: no cover
        return Path(os.path.abspath(sys.argv[0])).resolve().parent


def dskm_data_dir() -> Path:
    """可写数据目录：优先程序目录；不可写时退到 %APPDATA%。"""
    base = dskm_app_dir()
    try:
        probe = base / ".dskm_write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return base
    except Exception:
        pass
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    fallback = Path(appdata) / DSKM_APP_ID
    try:
        fallback.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return fallback


def dskm_env_path() -> Path:
    """.env 文件路径（优先程序目录已存在的那个）。"""
    for candidate in (dskm_app_dir() / ".env", dskm_data_dir() / ".env"):
        try:
            if candidate.exists():
                return candidate
        except Exception:
            continue
    return dskm_data_dir() / ".env"


def dskm_settings_path() -> Path:
    return dskm_data_dir() / "deepseek_key_manager_settings.json"


def dskm_meta_path() -> Path:
    return dskm_data_dir() / "keys_meta.json"


def dskm_log_dir() -> Path:
    d = dskm_data_dir() / "logs"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def dskm_icon() -> QIcon:
    """图标：检测 icon.ico 是否存在，存在则使用，不存在返回空 QIcon（不报错）。"""
    if getattr(dskm_icon, "_cache", None) is not None:
        return dskm_icon._cache  # type: ignore[attr-defined]
    icon = QIcon()
    candidates = [
        dskm_app_dir() / "icon.ico",
        dskm_app_dir() / "data" / "icon.ico",
        dskm_app_dir() / "resources" / "icon.ico",
    ]
    if dskm_is_compiled():
        candidates.insert(0, Path(sys.executable).resolve().parent / "icon.ico")
    for path in candidates:
        try:
            if path.exists() and path.stat().st_size > 0:
                loaded = QIcon(str(path))
                if not loaded.isNull() and not loaded.availableSizes() == []:
                    icon = loaded
                    break
        except Exception:
            continue
    try:
        dskm_icon._cache = icon  # type: ignore[attr-defined]
    except Exception:
        pass
    return icon


def dskm_now_text() -> str:
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def dskm_now_stamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def dskm_parse_time_text(text: str) -> Optional[_dt.datetime]:
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S"):
        try:
            return _dt.datetime.strptime((text or "").strip(), fmt)
        except Exception:
            continue
    return None


def dskm_human_seconds(seconds: float) -> str:
    try:
        seconds = float(seconds)
    except Exception:
        return ""
    if seconds < 0:
        return ""
    if seconds < 60:
        return f"{seconds:.0f} 秒"
    if seconds < 3600:
        return f"{seconds / 60:.1f} 分钟"
    if seconds < 86400:
        return f"{seconds / 3600:.1f} 小时"
    return f"{seconds / 86400:.1f} 天"


def dskm_key_uid(api_key: str) -> str:
    """由 API Key 派生的稳定标识（不泄露原文，用于关联元数据）。"""
    raw = (api_key or "").strip().encode("utf-8", "ignore")
    return hashlib.sha256(raw).hexdigest()[:16]


def dskm_format_key_display(
    api_key: str,
    mode: str = KEY_DISPLAY_HEAD_TAIL,
    head: int = 6,
    tail: int = 4,
    mask_char: str = "*",
    mask_length: int = 8,
) -> str:
    """按用户设置格式化 API Key 的显示文本（列表/导出/剪贴板/日志 的唯一来源）。

    mode：
      full       完整显示
      head_tail  前 N 位 + 掩码 + 后 M 位
      head_only  只显示前 N 位
      tail_only  只显示后 M 位
      all_mask   全部用掩码符号
    mask_length：0 表示“按被隐去的字符数显示等量掩码符号”，>0 表示固定这么多个
                 （固定长度不泄露 Key 真实长度，更安全）。
    """
    key = (api_key or "").strip()
    if not key:
        return ""
    char = (str(mask_char or "*").strip() or "*")[:1]
    mode = str(mode or KEY_DISPLAY_HEAD_TAIL)
    if mode not in KEY_DISPLAY_MODE_LABELS:
        mode = KEY_DISPLAY_HEAD_TAIL
    head = max(0, min(64, int(head or 0)))
    tail = max(0, min(64, int(tail or 0)))
    try:
        mask_length = max(0, min(64, int(mask_length or 0)))
    except Exception:
        mask_length = 8

    def _mask(n_hidden: int) -> str:
        count = mask_length if mask_length > 0 else max(3, min(64, int(n_hidden)))
        return char * count

    if mode == KEY_DISPLAY_FULL:
        return key

    if mode == KEY_DISPLAY_ALL_MASK:
        return _mask(len(key) if mask_length <= 0 else 0)

    if mode == KEY_DISPLAY_HEAD_ONLY:
        visible = min(head, max(0, len(key) - 2))
        return key[:visible] + _mask(len(key) - visible)

    if mode == KEY_DISPLAY_TAIL_ONLY:
        visible = min(tail, max(0, len(key) - 2))
        return _mask(len(key) - visible) + (key[len(key) - visible:] if visible else "")

    # head_tail：保证至少留下 2 个字符被掩码，装不下就退化为全部掩码
    head = min(head, max(0, len(key) - 2))
    tail = min(tail, max(0, len(key) - head - 2))
    head_part = key[:head]
    tail_part = key[len(key) - tail:] if tail > 0 else ""
    hidden = len(key) - len(head_part) - len(tail_part)
    if hidden <= 0:
        return _mask(len(key))
    return f"{head_part}{_mask(hidden)}{tail_part}"


def dskm_mask_key(api_key: str, head: int = 6, tail: int = 4, full_hide: bool = False,
                  mask_char: str = "*", mask_length: int = 8) -> str:
    """掩码显示 API Key（`dskm_format_key_display` 的简化封装，保留旧签名兼容）。

    默认与程序默认设置一致：前 6 位 + 8 个掩码符号 + 后 4 位（固定 8 个，不泄露 Key 长度）。
    """
    mode = KEY_DISPLAY_ALL_MASK if full_hide else KEY_DISPLAY_HEAD_TAIL
    return dskm_format_key_display(api_key, mode, head, tail, mask_char, mask_length)


def dskm_normalize_key(api_key: str) -> str:
    """清理用户粘贴的 Key（去空白、去引号、去前缀）。"""
    key = (api_key or "").strip()
    key = key.strip("\"'")
    key = key.replace("\r", "").replace("\n", "").replace("\t", "").replace(" ", "")
    for prefix in ("Bearer", "bearer", "Authorization:"):
        if key.startswith(prefix):
            key = key[len(prefix):].strip(": ").strip()
    return key


_KEY_PATTERN = re.compile(r"(sk-[A-Za-z0-9_\-]{6,})")

# 兜底抽取时要排除的“看起来不是密钥”的值：URL / IP / 路径 / 邮箱 / 纯数字
_NOT_KEY_PATTERN = re.compile(r"(://)|[@]|[/\\]|[.]|\s")


def dskm_looks_like_value_key(text: str) -> bool:
    """判断一个变量值是否像 API Key（用于兜底抽取时排除配置项）。"""
    value = dskm_normalize_key(text)
    if len(value) < 12:
        return False
    if dskm_is_ascii_digits(value):
        return False
    if _NOT_KEY_PATTERN.search(value):
        return False
    return True


def dskm_extract_keys(text: str) -> List[str]:
    """从任意文本中抽取 API Key（兼容 .env / txt / json / 剪贴板 粘贴）。"""
    found: List[str] = []
    seen = set()
    raw = text or ""
    # 1) 优先按 sk- 前缀抓取
    for match in _KEY_PATTERN.finditer(raw):
        key = dskm_normalize_key(match.group(1))
        if key and key not in seen:
            seen.add(key)
            found.append(key)
    # 2) 兜底：按行 / 逗号 / 分号 / 竖线 切分（兼容用户粘贴非 sk- 前缀的自定义网关 Key）
    #    这里必须严格过滤，否则 .env 里的 DEEPSEEK_API_BASE=https://...、DB_HOST=127.0.0.1
    #    会被误当成密钥（实测踩到过）
    for sep_line in raw.replace(",", "\n").replace(";", "\n").replace("|", "\n").splitlines():
        line = sep_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            line = line.split("=", 1)[1].strip()
        key = dskm_normalize_key(line)
        if not key or key.lower() in ("dskm_none", "none", "null", "true", "false"):
            continue
        if not dskm_looks_like_value_key(key):
            continue
        if key not in seen:
            seen.add(key)
            found.append(key)
    return found


def dskm_looks_like_api_key(text: str) -> bool:
    key = dskm_normalize_key(text)
    return len(key) >= 12 and (" " not in key)


def dskm_to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        try:
            return float(Decimal(text))
        except (InvalidOperation, Exception):
            return None


def dskm_to_decimal(value: Any) -> Optional[Decimal]:
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        return Decimal(text)
    except Exception:
        return None


def dskm_format_amount(value: Any, precision: int = 2) -> str:
    """金额格式化：保留原值精度但按设定精度补齐，空值返回空串。"""
    if value is None:
        return ""
    dec = dskm_to_decimal(value)
    if dec is None:
        return str(value)
    precision = max(0, min(6, int(precision or 0)))
    quant = Decimal(1).scaleb(-precision)
    try:
        dec = dec.quantize(quant)
    except Exception:
        pass
    return f"{dec:f}"


def dskm_atomic_write_text(path: Path, text: str, encoding: str = "utf-8") -> None:
    """原子写入：先写临时文件再 os.replace，避免断电/崩溃导致文件损坏。"""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    data = text.encode(encoding)
    with open(tmp, "wb") as fp:
        fp.write(data)
        fp.flush()
        try:
            os.fsync(fp.fileno())
        except Exception:
            pass
    os.replace(tmp, path)


def dskm_atomic_write_json(path: Path, payload: Any) -> None:
    dskm_atomic_write_text(path, json.dumps(payload, ensure_ascii=False, indent=2))


def dskm_safe_read_json(path: Path, default: Any = None) -> Any:
    try:
        with open(path, "r", encoding="utf-8") as fp:
            return json.load(fp)
    except Exception:
        return default


def dskm_is_ascii_digits(text: str) -> bool:
    """严格判断 ASCII 数字串（禁用 str.isdigit：全角/上标也会判真）。"""
    s = str(text)
    if not s:
        return False
    return all("0" <= ch <= "9" for ch in s)


def dskm_safe_filename(name: str, fallback: str = "file") -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(name or "")).strip(" .")
    return cleaned or fallback


def dskm_elide(text: str, limit: int = 120) -> str:
    text = str(text or "")
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"


# ---- 屏幕安全区：窗口几何（沿用既有项目验证过的零漂移做法）----------------


def dskm_screen_available_rects() -> List[QRect]:
    rects: List[QRect] = []
    try:
        for screen in QGuiApplication.screens():
            rects.append(screen.availableGeometry())
    except Exception:
        pass
    if not rects:
        try:
            rects.append(QGuiApplication.primaryScreen().availableGeometry())
        except Exception:
            pass
    if not rects:
        rects.append(QRect(0, 0, 1920, 1080))
    return rects


def dskm_rect_inside(a: QRect, b: QRect, margin: int = 0) -> bool:
    return (
        a.x() >= b.x() - margin
        and a.y() >= b.y() - margin
        and a.right() <= b.right() + margin
        and a.bottom() <= b.bottom() + margin
    )


def dskm_safe_window_geometry(
    saved: Any,
    default_size: Sequence[int] = (1420, 860),
    min_visible_w: int = 160,
    min_visible_h: int = 60,
    always_center: bool = False,
) -> List[int]:
    """把保存的窗口几何夹到屏幕安全区内，避免窗口跑到屏幕外抓不住。"""
    rects = dskm_screen_available_rects()
    primary = rects[0]
    dw = int(default_size[0])
    dh = int(default_size[1])
    try:
        if isinstance(saved, (list, tuple)) and len(saved) >= 4:
            x, y, w, h = (int(saved[0]), int(saved[1]), int(saved[2]), int(saved[3]))
        else:
            raise ValueError
    except Exception:
        w, h = dw, dh
        x = primary.x() + max(0, (primary.width() - w) // 2)
        y = primary.y() + max(0, (primary.height() - h) // 3)
        return [x, y, w, h]

    w = max(760, min(w, max(primary.width(), 760)))
    h = max(520, min(h, max(primary.height(), 520)))

    if always_center:
        x = primary.x() + max(0, (primary.width() - w) // 2)
        y = primary.y() + max(0, (primary.height() - h) // 3)
        return [x, y, w, h]

    for rect in rects:
        probe = QRect(x, y, w, h)
        visible_w = min(probe.right(), rect.right()) - max(probe.x(), rect.x())
        visible_h = min(probe.bottom(), rect.bottom()) - max(probe.y(), rect.y())
        if visible_w >= min(min_visible_w, w) and visible_h >= min(min_visible_h, h):
            if probe.x() < rect.x():
                x = rect.x()
            if probe.y() < rect.y():
                y = rect.y()
            if probe.right() > rect.right():
                x = max(rect.x(), rect.right() - w + 1)
            if probe.bottom() > rect.bottom():
                y = max(rect.y(), rect.bottom() - h + 1)
            return [x, y, w, h]
    x = primary.x() + max(0, (primary.width() - w) // 2)
    y = primary.y() + max(0, (primary.height() - h) // 3)
    return [x, y, w, h]


def dskm_place_window(window: QWidget, geometry: Sequence[int]) -> None:
    """摆放窗口：resize + move（move 与 x()/y() 同坐标系，避免每启动一次往上顶）。"""
    try:
        x, y, w, h = (int(geometry[0]), int(geometry[1]), int(geometry[2]), int(geometry[3]))
    except Exception:
        return
    try:
        window.resize(w, h)
        window.move(x, y)
    except Exception:
        pass


def dskm_window_geometry_for_save(window: QWidget, default: Sequence[int]) -> List[int]:
    """保存窗口几何（最大化时保存 normalGeometry，避免记录 -8/-31 残影）。"""
    try:
        if window.isMaximized():
            geo = window.normalGeometry()
        else:
            geo = window.frameGeometry()
            if not geo.isValid():
                geo = QRect(int(window.x()), int(window.y()), window.width(), window.height())
        return [geo.x(), geo.y(), max(760, geo.width()), max(520, geo.height())]
    except Exception:
        return list(default)


def dskm_ensure_window_on_screen(window: QWidget) -> None:
    """show 之后按真实 frameGeometry 兜底，确保窗口可达。"""
    try:
        geo = window.frameGeometry()
        if not geo.isValid():
            return
        for rect in dskm_screen_available_rects():
            visible_w = min(geo.right(), rect.right()) - max(geo.x(), rect.x())
            visible_h = min(geo.bottom(), rect.bottom()) - max(geo.y(), rect.y())
            if visible_w >= 120 and visible_h >= 40:
                return
        primary = dskm_screen_available_rects()[0]
        w = min(geo.width(), primary.width())
        h = min(geo.height(), primary.height())
        x = primary.x() + max(0, (primary.width() - w) // 2)
        y = primary.y() + max(0, (primary.height() - h) // 3)
        window.move(x, y)
    except Exception:
        pass


def dskm_best_ui_font(size: int = 9) -> QFont:
    """挑选一个中文显示效果好的界面字体。"""
    families = set(QFontDatabase.families())
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "Segoe UI", "Noto Sans CJK SC"):
        if name in families:
            return QFont(name, size)
    return QFont("", size)


def dskm_mono_font(size: int = 9) -> QFont:
    families = set(QFontDatabase.families())
    for name in ("Consolas", "Cascadia Mono", "DejaVu Sans Mono", "Courier New"):
        if name in families:
            return QFont(name, size)
    return QFont("", size)


# ============================================================================
#  区域 3：.env 存储（API Key / 数据库密码 一律放这里，绝不硬编码）
# ============================================================================


class DSKMEnvStore:
    """极简 .env 读写器。

    设计目标：
      • **保留原有注释、顺序、未知变量**（用 python-dotenv 的 set_key 会重排）
      • 支持 `KEY=value`、`KEY="value"`、`KEY='value'`
      • 支持 `KEY=sk-a,sk-b` 与 `KEY_1=sk-a` / `KEY_2=sk-b` 两种密钥写法
      • 原子写入（tmp + os.replace），避免密钥文件被写坏
    """

    KEY_API_KEYS = "***已脱敏***-apikey"
    KEY_API_BASE = "DEEPSEEK_API_BASE"
    KEY_TIMEOUT = "DEEPSEEK_TIMEOUT"
    KEY_MAX_RETRY = "DEEPSEEK_MAX_RETRY"
    KEY_CONCURRENCY = "DEEPSEEK_CONCURRENCY"
    KEY_DB_PASSWORD = "DB_PASSWORD"

    DEFAULT_TEMPLATE = (
        "# ============================================================\n"
        "# DeepSeek 密钥余额管理器 —— 敏感配置（请勿提交到公开仓库）\n"
        "# 本文件由程序自动维护，手工编辑后再次启动即可生效。\n"
        "# ============================================================\n"
        "\n"
        "# ---- DeepSeek 接口 ----\n"
        "DEEPSEEK_API_BASE=https://api.deepseek.com\n"
        "# 多个 Key 用英文逗号分隔，没有数量上限\n"
        "DEEPSEEK_API_KEYS=\n"
        "DEEPSEEK_TIMEOUT=20\n"
        "DEEPSEEK_MAX_RETRY=2\n"
        "DEEPSEEK_CONCURRENCY=5\n"
        "\n"
        "# ---- 数据库（当前版本预留，供后续扩展使用）----\n"
        "DB_HOST=127.0.0.1\n"
        "DB_PORT=3306\n"
        "DB_USER=\n"
        "DB_PASSWORD=\n"
        "DB_NAME=\n"
    )

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else dskm_env_path()
        self._lock = threading.RLock()

    # ---- 基础读写 -------------------------------------------------------
    def exists(self) -> bool:
        try:
            return self.path.exists()
        except Exception:
            return False

    def read_text(self) -> str:
        with self._lock:
            try:
                with open(self.path, "r", encoding="utf-8-sig") as fp:
                    return fp.read()
            except Exception:
                return ""

    def write_text(self, text: str) -> None:
        with self._lock:
            try:
                self.path.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            dskm_atomic_write_text(self.path, text, "utf-8")

    def ensure_exists(self) -> None:
        """首次运行时自动生成 .env 模板（不覆盖已有文件）。"""
        if not self.exists():
            self.write_text(self.DEFAULT_TEMPLATE)

    # ---- 解析 -----------------------------------------------------------
    @staticmethod
    def parse_pairs(text: str) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for raw_line in (text or "").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.lower().startswith("export "):
                line = line[7:].strip()
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                value = value[1:-1]
            if key:
                result[key] = value
        return result

    def as_dict(self) -> Dict[str, str]:
        return self.parse_pairs(self.read_text())

    def get(self, key: str, default: str = "") -> str:
        return self.as_dict().get(key, default)

    def get_int(self, key: str, default: int) -> int:
        try:
            return int(str(self.get(key, "")).strip())
        except Exception:
            return default

    def get_bool(self, key: str, default: bool) -> bool:
        raw = str(self.get(key, "")).strip().lower()
        if raw in ("1", "true", "yes", "on", "y"):
            return True
        if raw in ("0", "false", "no", "off", "n"):
            return False
        return default

    # ---- 写入（保留注释与顺序）-----------------------------------------
    def set_values(self, mapping: Dict[str, str], remove_keys: Iterable[str] = ()) -> None:
        with self._lock:
            text = self.read_text()
            if not text.strip():
                text = self.DEFAULT_TEMPLATE
            remove_set = {str(k) for k in remove_keys}
            lines = text.splitlines()
            written: Dict[str, str] = {}
            out_lines: List[str] = []
            for raw_line in lines:
                stripped = raw_line.strip()
                if not stripped or stripped.startswith("#") or "=" not in stripped:
                    out_lines.append(raw_line)
                    continue
                key = stripped.split("=", 1)[0].strip()
                if key in remove_set:
                    continue
                if key in mapping:
                    out_lines.append(f"{key}={mapping[key]}")
                    written[key] = mapping[key]
                else:
                    out_lines.append(raw_line)
            pending = {k: v for k, v in mapping.items() if k not in written}
            if pending:
                while out_lines and not out_lines[-1].strip():
                    out_lines.pop()
                out_lines.append("")
                out_lines.append("# ---- 由程序自动写入 ----")
                for key, value in pending.items():
                    out_lines.append(f"{key}={value}")
            new_text = "\n".join(out_lines).rstrip("\n") + "\n"
            self.write_text(new_text)

    def remove_keys(self, keys: Iterable[str]) -> None:
        with self._lock:
            remove_set = {str(k) for k in keys}
            if not remove_set:
                return
            text = self.read_text()
            out_lines: List[str] = []
            for raw_line in text.splitlines():
                stripped = raw_line.strip()
                if stripped and not stripped.startswith("#") and "=" in stripped:
                    key = stripped.split("=", 1)[0].strip()
                    if key in remove_set:
                        continue
                out_lines.append(raw_line)
            self.write_text("\n".join(out_lines).rstrip("\n") + "\n")

    # ---- API Key 专用 ---------------------------------------------------
    def get_api_keys(self) -> List[str]:
        """读取所有 API Key：兼容逗号分隔与 `DEEPSEEK_API_KEY_1..N` 两种写法。"""
        data = self.as_dict()
        keys: List[str] = []
        seen = set()

        def _push(raw: str) -> None:
            for item in re.split(r"[,\n;|]+", str(raw or "")):
                key = dskm_normalize_key(item)
                if key and key not in seen:
                    seen.add(key)
                    keys.append(key)

        _push(data.get(self.KEY_API_KEYS, ""))
        numbered: List[Tuple[int, str]] = []
        for name, value in data.items():
            if re.fullmatch(r"DEEPSEEK_API_KEY_\d+", name):
                try:
                    idx = int(name.rsplit("_", 1)[1])
                except Exception:
                    idx = 9999
                numbered.append((idx, value))
            elif name in ("DEEPSEEK_API_KEY", "DEEPSEEK_KEY"):
                numbered.append((0, value))
        numbered.sort(key=lambda item: item[0])
        for _, value in numbered:
            _push(value)
        # 兜底：只从「变量名看起来就是密钥」的配置项里取（绝不整文件乱抽）
        if not keys:
            for name, value in data.items():
                upper = str(name).upper()
                compact = upper.replace("_", "")
                if ("APIKEY" in compact or "TOKEN" in upper or "SECRET" in upper
                        or upper.endswith("KEY") or upper.endswith("KEYS")):
                    _push(value)
        return keys

    def set_api_keys(self, keys: Sequence[str]) -> None:
        cleaned: List[str] = []
        seen = set()
        for item in keys:
            key = dskm_normalize_key(item)
            if key and key not in seen:
                seen.add(key)
                cleaned.append(key)
        remove = [name for name in self.as_dict() if re.fullmatch(r"DEEPSEEK_API_KEY(_\d+)?", name)]
        self.set_values({self.KEY_API_KEYS: ",".join(cleaned)}, remove_keys=remove)

    def get_api_base(self) -> str:
        base = (self.get(self.KEY_API_BASE, DSKM_DEFAULT_API_BASE) or "").strip()
        if not base:
            base = DSKM_DEFAULT_API_BASE
        return base.rstrip("/")

    def get_network_options(self) -> Dict[str, Any]:
        return {
            "api_base": self.get_api_base(),
            "timeout": max(3, min(300, self.get_int(self.KEY_TIMEOUT, 20))),
            "max_retry": max(0, min(10, self.get_int(self.KEY_MAX_RETRY, 2))),
            "concurrency": max(1, min(64, self.get_int(self.KEY_CONCURRENCY, 5))),
        }

    def set_network_options(
        self, api_base: Optional[str] = None, timeout: Optional[int] = None,
        max_retry: Optional[int] = None, concurrency: Optional[int] = None,
    ) -> None:
        mapping: Dict[str, str] = {}
        if api_base is not None:
            mapping[self.KEY_API_BASE] = str(api_base).strip().rstrip("/") or DSKM_DEFAULT_API_BASE
        if timeout is not None:
            mapping[self.KEY_TIMEOUT] = str(int(timeout))
        if max_retry is not None:
            mapping[self.KEY_MAX_RETRY] = str(int(max_retry))
        if concurrency is not None:
            mapping[self.KEY_CONCURRENCY] = str(int(concurrency))
        if mapping:
            self.set_values(mapping)


# ============================================================================
#  区域 4：程序设置（JSON 原子写 + 防抖自动保存）
# ============================================================================


class DSKMSettingsStore:
    """设置存储：内存字典 + 原子落盘。所有设置改动都应即时保存，避免重启丢配置。"""

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else dskm_settings_path()
        self._lock = threading.RLock()
        self._data: Dict[str, Any] = dict(DEFAULT_SETTINGS)
        self._dirty = False
        self.load()

    def load(self) -> Dict[str, Any]:
        raw = dskm_safe_read_json(self.path, {})
        data = dict(DEFAULT_SETTINGS)
        if isinstance(raw, dict):
            for key, value in raw.items():
                data[key] = value
        # 老版本兼容：只保留已知键 + 未知键也保留（避免丢配置）
        self._migrate_legacy_keys(raw, data)
        self._data = data
        return self._data

    @staticmethod
    def _migrate_legacy_keys(raw: Any, data: Dict[str, Any]) -> None:
        """把 v1.0.0 的 mask_api_key / mask_full_hide 迁移成 v1.1.0 的 key_display_mode。

        只在磁盘上还没写过新键时迁移一次，避免覆盖用户已经选好的显示方式。
        """
        if not isinstance(raw, dict) or SET_KEY_KEY_DISPLAY_MODE in raw:
            return
        if SET_KEY_MASK_API_KEY not in raw and SET_KEY_MASK_FULL not in raw:
            return
        mask_on = raw.get(SET_KEY_MASK_API_KEY, True)
        full_hide = raw.get(SET_KEY_MASK_FULL, False)
        if mask_on is False:
            data[SET_KEY_KEY_DISPLAY_MODE] = KEY_DISPLAY_FULL
        elif full_hide is True:
            data[SET_KEY_KEY_DISPLAY_MODE] = KEY_DISPLAY_ALL_MASK
        else:
            data[SET_KEY_KEY_DISPLAY_MODE] = KEY_DISPLAY_HEAD_TAIL

    @property
    def data(self) -> Dict[str, Any]:
        return self._data

    def get(self, key: str, default: Any = None) -> Any:
        if key in self._data:
            return self._data[key]
        if key in DEFAULT_SETTINGS:
            return DEFAULT_SETTINGS[key]
        return default

    def get_bool(self, key: str, default: bool = False) -> bool:
        value = self.get(key, default)
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        text = str(value).strip().lower()
        if text in ("1", "true", "yes", "on"):
            return True
        if text in ("0", "false", "no", "off", ""):
            return False
        return bool(default)

    def get_int(self, key: str, default: int = 0) -> int:
        try:
            return int(self.get(key, default))
        except Exception:
            return int(default)

    def get_float(self, key: str, default: float = 0.0) -> float:
        try:
            return float(self.get(key, default))
        except Exception:
            return float(default)

    def get_str(self, key: str, default: str = "") -> str:
        value = self.get(key, default)
        if value is None:
            return default
        return str(value)

    def get_list(self, key: str, default: Optional[list] = None) -> list:
        value = self.get(key, default if default is not None else [])
        if isinstance(value, list):
            return list(value)
        if isinstance(value, (tuple, set)):
            return list(value)
        return list(default or [])

    def get_dict(self, key: str, default: Optional[dict] = None) -> dict:
        value = self.get(key, default if default is not None else {})
        if isinstance(value, dict):
            return dict(value)
        return dict(default or {})

    def set(self, key: str, value: Any, save_now: bool = False) -> None:
        with self._lock:
            if self._data.get(key) == value:
                if save_now:
                    self.save()
                return
            self._data[key] = value
            self._dirty = True
        if save_now:
            self.save()

    def update(self, mapping: Dict[str, Any], save_now: bool = False) -> None:
        if not mapping:
            return
        with self._lock:
            self._data.update(mapping)
            self._dirty = True
        if save_now:
            self.save()

    def save(self) -> bool:
        with self._lock:
            payload = dict(self._data)
            self._dirty = False
        try:
            dskm_atomic_write_json(self.path, payload)
            return True
        except Exception:
            return False

    def is_dirty(self) -> bool:
        return self._dirty


# ============================================================================
#  区域 5：密钥元数据存储（别名/备注/启用状态/历史余额）
# ============================================================================


class DSKMKeyMetaStore:
    """按 uid 保存每个密钥的元数据（**不保存 Key 原文**，Key 只存在于 .env）。"""

    FORMAT_VERSION = 1

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else dskm_meta_path()
        self._lock = threading.RLock()
        self._records: Dict[str, Dict[str, Any]] = {}
        self.load()

    def load(self) -> Dict[str, Dict[str, Any]]:
        raw = dskm_safe_read_json(self.path, {})
        records: Dict[str, Dict[str, Any]] = {}
        if isinstance(raw, dict):
            items = raw.get("records", raw)
            if isinstance(items, dict):
                for uid, payload in items.items():
                    if isinstance(payload, dict):
                        records[str(uid)] = payload
            elif isinstance(items, list):
                for payload in items:
                    if isinstance(payload, dict) and payload.get("uid"):
                        records[str(payload["uid"])] = payload
        self._records = records
        return records

    def save(self) -> bool:
        with self._lock:
            payload = {
                "format_version": self.FORMAT_VERSION,
                "updated_at": dskm_now_text(),
                "records": {uid: dict(rec) for uid, rec in self._records.items()},
            }
        try:
            dskm_atomic_write_json(self.path, payload)
            return True
        except Exception:
            return False

    def get(self, uid: str) -> Dict[str, Any]:
        with self._lock:
            return dict(self._records.get(uid, {}))

    def set(self, uid: str, payload: Dict[str, Any], save: bool = True) -> None:
        if not uid:
            return
        with self._lock:
            merged = dict(self._records.get(uid, {}))
            merged.update(payload)
            merged["uid"] = uid
            self._records[uid] = merged
        if save:
            self.save()

    def remove(self, uid: str, save: bool = True) -> None:
        with self._lock:
            existed = self._records.pop(uid, None) is not None
        if existed and save:
            self.save()

    def prune(self, valid_uids: Iterable[str], save: bool = True) -> int:
        """清理已不在 .env 里的残留元数据（默认保留历史，仅在显式调用时清理）。"""
        valid = {str(u) for u in valid_uids}
        with self._lock:
            stale = [uid for uid in self._records if uid not in valid]
            for uid in stale:
                self._records.pop(uid, None)
        if stale and save:
            self.save()
        return len(stale)


# ============================================================================
#  区域 6：日志（内存信号 + 文件轮转）
# ============================================================================


class DSKMLogBroker(QObject):
    """全局日志总线：界面订阅 signal，文件写入在内部完成。"""

    message_logged = Signal(str, str)  # level, text

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._file_lock = threading.RLock()
        self._max_file_bytes = 2 * 1024 * 1024
        self._keep_files = 5
        self._min_level = "debug"

    def set_min_level(self, level: str) -> None:
        self._min_level = str(level or "debug").lower()

    def _level_enabled(self, level: str) -> bool:
        order = {"debug": 0, "info": 1, "success": 1, "warn": 2, "error": 3}
        return order.get(str(level).lower(), 1) >= order.get(self._min_level, 0)

    def log(self, level: str, text: str) -> None:
        level = str(level or "info").lower()
        if not self._level_enabled(level):
            return
        stamp = _dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{stamp}] [{DSKM_LOG_LEVELS.get(level, level)}] {text}"
        self._write_file(level, line)
        try:
            self.message_logged.emit(level, line)
        except Exception:
            pass

    def debug(self, text: str) -> None:
        self.log("debug", text)

    def info(self, text: str) -> None:
        self.log("info", text)

    def success(self, text: str) -> None:
        self.log("success", text)

    def warn(self, text: str) -> None:
        self.log("warn", text)

    def error(self, text: str) -> None:
        self.log("error", text)

    def exception(self, text: str) -> None:
        self.log("error", f"{text}\n{traceback.format_exc()}")

    # ---- 文件写入 -------------------------------------------------------
    def _current_log_file(self) -> Path:
        return dskm_log_dir() / f"app_{_dt.datetime.now().strftime('%Y%m%d')}.log"

    def _rotate_if_needed(self, path: Path) -> None:
        try:
            if not path.exists() or path.stat().st_size < self._max_file_bytes:
                return
            rotated = path.with_name(f"{path.stem}_{dskm_now_stamp()}{path.suffix}")
            os.replace(path, rotated)
            files = sorted(dskm_log_dir().glob("app_*.log"), key=lambda p: p.stat().st_mtime)
            for old in files[: max(0, len(files) - self._keep_files)]:
                try:
                    old.unlink()
                except Exception:
                    pass
        except Exception:
            pass

    def _write_file(self, level: str, line: str) -> None:
        with self._file_lock:
            try:
                path = self._current_log_file()
                self._rotate_if_needed(path)
                with open(path, "a", encoding="utf-8") as fp:
                    fp.write(line + "\n")
            except Exception:
                pass


# 全局日志总线（模块级单例，供各处直接使用）
DSKM_LOG = DSKMLogBroker()


# ============================================================================
#  区域 7：数据模型
# ============================================================================


@dataclass
class DSKMBalanceInfo:
    """官方接口 balance_infos 数组中的一条余额明细。"""

    currency: str = ""
    total_balance: str = ""
    granted_balance: str = ""
    topped_up_balance: str = ""

    @classmethod
    def from_api(cls, payload: Dict[str, Any]) -> "DSKMBalanceInfo":
        payload = payload or {}
        return cls(
            currency=str(payload.get(DSKM_FIELD_CURRENCY, "") or ""),
            total_balance=str(payload.get(DSKM_FIELD_TOTAL_BALANCE, "") or ""),
            granted_balance=str(payload.get(DSKM_FIELD_GRANTED_BALANCE, "") or ""),
            topped_up_balance=str(payload.get(DSKM_FIELD_TOPPED_UP_BALANCE, "") or ""),
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            DSKM_FIELD_CURRENCY: self.currency,
            DSKM_FIELD_TOTAL_BALANCE: self.total_balance,
            DSKM_FIELD_GRANTED_BALANCE: self.granted_balance,
            DSKM_FIELD_TOPPED_UP_BALANCE: self.topped_up_balance,
        }

    def display_text(self, precision: int = 2) -> str:
        cur = self.currency or "?"
        total = dskm_format_amount(self.total_balance, precision) or "-"
        granted = dskm_format_amount(self.granted_balance, precision) or "-"
        topped = dskm_format_amount(self.topped_up_balance, precision) or "-"
        return f"{cur} 总 {total}（赠金 {granted} / 充值 {topped}）"

    def currency_symbol(self) -> str:
        return {"CNY": "¥", "USD": "$", "EUR": "€", "JPY": "¥"}.get(self.currency.upper(), "")


@dataclass
class DSKMQueryResult:
    """一次余额查询的结果。"""

    uid: str = ""
    ok: bool = False
    is_available: Optional[bool] = None
    balance_infos: List[DSKMBalanceInfo] = field(default_factory=list)
    http_status: int = 0
    error_message: str = ""
    error_kind: str = ""
    elapsed_ms: int = 0
    checked_at: str = ""
    raw: Dict[str, Any] = field(default_factory=dict)

    @property
    def primary(self) -> Optional[DSKMBalanceInfo]:
        return self.balance_infos[0] if self.balance_infos else None


@dataclass
class DSKMKeyRecord:
    """一个 API Key 的完整运行时记录（Key 原文来自 .env）。"""

    api_key: str = ""
    uid: str = ""
    alias: str = ""
    note: str = ""
    enabled: bool = True
    created_at: str = ""

    last_checked: str = ""
    last_status: str = STATUS_PENDING
    progress: int = 0
    is_available: Optional[bool] = None
    balance_infos: List[DSKMBalanceInfo] = field(default_factory=list)
    http_status: int = 0
    error_message: str = ""
    error_kind: str = ""
    elapsed_ms: int = 0
    check_count: int = 0
    fail_count: int = 0
    success_count: int = 0
    history: List[Dict[str, Any]] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.uid:
            self.uid = dskm_key_uid(self.api_key)
        if not self.created_at:
            self.created_at = dskm_now_text()

    # ---- 派生属性 -------------------------------------------------------
    @property
    def display_alias(self) -> str:
        return (self.alias or "").strip() or "（未命名）"

    @property
    def masked_key(self) -> str:
        """默认掩码文本（前 6 + `*` + 后 4）；列表/导出请用
        `DSKMKeyTableModel.format_key_text()` 以便跟随用户设置。"""
        return dskm_mask_key(self.api_key)

    @property
    def is_disabled(self) -> bool:
        return not self.enabled

    @property
    def status_text(self) -> str:
        if not self.enabled:
            return STATUS_TEXT[STATUS_DISABLED]
        if self.last_status == STATUS_OK and self.ok_flag:
            return "可用" if self.is_available else "不可用（账户被停用）"
        return STATUS_TEXT.get(self.last_status, self.last_status or "待查询")

    @property
    def ok_flag(self) -> bool:
        return self.last_status == STATUS_OK

    @property
    def primary_balance(self) -> Optional[DSKMBalanceInfo]:
        return self.balance_infos[0] if self.balance_infos else None

    @property
    def total_balance_text(self) -> str:
        if not self.balance_infos:
            return ""
        infos = [i for i in self.balance_infos if i.total_balance not in ("", None)]
        if not infos:
            return ""
        if len(infos) == 1:
            return dskm_format_amount(infos[0].total_balance)
        return " + ".join(f"{dskm_format_amount(i.total_balance)} {i.currency}" for i in infos)

    @property
    def currency_text(self) -> str:
        if not self.balance_infos:
            return ""
        return " / ".join(i.currency for i in self.balance_infos if i.currency)

    @property
    def granted_text(self) -> str:
        if not self.balance_infos:
            return ""
        return " + ".join(
            dskm_format_amount(i.granted_balance) for i in self.balance_infos
        )

    @property
    def topped_up_text(self) -> str:
        if not self.balance_infos:
            return ""
        return " + ".join(
            dskm_format_amount(i.topped_up_balance) for i in self.balance_infos
        )

    @property
    def balance_detail_text(self) -> str:
        if not self.balance_infos:
            return ""
        return " ｜ ".join(i.display_text() for i in self.balance_infos)

    def total_amount_of(self, currency: str) -> Optional[Decimal]:
        for info in self.balance_infos:
            if info.currency.upper() == str(currency).upper():
                return dskm_to_decimal(info.total_balance)
        return None

    def has_balance(self) -> bool:
        for info in self.balance_infos:
            value = dskm_to_decimal(info.total_balance)
            if value is not None and value > 0:
                return True
        return False

    # ---- 序列化 ---------------------------------------------------------
    def to_meta_dict(self) -> Dict[str, Any]:
        """写入 keys_meta.json 的内容（**绝不含 api_key**）。"""
        return {
            "uid": self.uid,
            "alias": self.alias,
            "note": self.note,
            "enabled": bool(self.enabled),
            "created_at": self.created_at,
            "last_checked": self.last_checked,
            "last_status": self.last_status,
            "is_available": self.is_available,
            "balance_infos": [i.to_dict() for i in self.balance_infos],
            "http_status": self.http_status,
            "error_message": self.error_message,
            "error_kind": self.error_kind,
            "elapsed_ms": self.elapsed_ms,
            "check_count": self.check_count,
            "fail_count": self.fail_count,
            "success_count": self.success_count,
            "history": list(self.history),
        }

    def apply_meta_dict(self, payload: Dict[str, Any]) -> None:
        if not payload:
            return
        self.alias = str(payload.get("alias", self.alias) or "")
        self.note = str(payload.get("note", self.note) or "")
        self.enabled = bool(payload.get("enabled", self.enabled))
        self.created_at = str(payload.get("created_at", self.created_at) or self.created_at)
        self.last_checked = str(payload.get("last_checked", "") or "")
        status = str(payload.get("last_status", STATUS_PENDING) or STATUS_PENDING)
        self.last_status = status if status in STATUS_TEXT else STATUS_PENDING
        available = payload.get("is_available", None)
        self.is_available = None if available is None else bool(available)
        infos = payload.get("balance_infos", [])
        self.balance_infos = [
            DSKMBalanceInfo.from_api(item) for item in infos if isinstance(item, dict)
        ]
        self.http_status = int(payload.get("http_status", 0) or 0)
        self.error_message = str(payload.get("error_message", "") or "")
        self.error_kind = str(payload.get("error_kind", "") or "")
        self.elapsed_ms = int(payload.get("elapsed_ms", 0) or 0)
        self.check_count = int(payload.get("check_count", 0) or 0)
        self.fail_count = int(payload.get("fail_count", 0) or 0)
        self.success_count = int(payload.get("success_count", 0) or 0)
        history = payload.get("history", [])
        self.history = [h for h in history if isinstance(h, dict)] if isinstance(history, list) else []
        self.progress = 100 if self.last_status in (STATUS_OK, STATUS_FAIL) else 0

    def apply_result(self, result: DSKMQueryResult, history_limit: int = 200) -> None:
        self.last_checked = result.checked_at or dskm_now_text()
        self.elapsed_ms = result.elapsed_ms
        self.http_status = result.http_status
        self.check_count += 1
        if result.ok:
            self.last_status = STATUS_OK
            self.success_count += 1
            self.is_available = result.is_available
            self.balance_infos = list(result.balance_infos)
            self.error_message = ""
            self.error_kind = ""
            primary = result.primary
            if primary is not None and history_limit > 0:
                amount = dskm_to_decimal(primary.total_balance)
                self.history.append({
                    "time": self.last_checked,
                    "currency": primary.currency,
                    "total": str(primary.total_balance),
                    "granted": str(primary.granted_balance),
                    "topped_up": str(primary.topped_up_balance),
                    "amount": float(amount) if amount is not None else None,
                })
                if len(self.history) > history_limit:
                    self.history = self.history[-history_limit:]
        else:
            self.last_status = STATUS_FAIL
            self.fail_count += 1
            self.error_message = result.error_message
            self.error_kind = result.error_kind
        self.progress = 100


# ============================================================================
#  区域 8：DeepSeek 接口客户端
# ============================================================================


class DSKMDeepSeekClient:
    """DeepSeek 余额接口客户端。

    接口：GET {api_base}/user/balance
    请求头：Authorization: Bearer <API Key> / Accept: application/json
    返回：{"is_available": true, "balance_infos": [{"currency": "CNY",
           "total_balance": "110.00", "granted_balance": "10.00",
           "topped_up_balance": "100.00"}]}
    """

    def __init__(
        self,
        api_base: str = DSKM_DEFAULT_API_BASE,
        timeout: int = 20,
        max_retry: int = 2,
        verify_ssl: bool = True,
        proxy_url: str = "",
        user_agent: str = DSKM_USER_AGENT,
    ) -> None:
        self.api_base = (api_base or DSKM_DEFAULT_API_BASE).rstrip("/")
        self.timeout = max(3, int(timeout or 20))
        self.max_retry = max(0, int(max_retry or 0))
        self.verify_ssl = bool(verify_ssl)
        self.proxy_url = (proxy_url or "").strip()
        self.user_agent = user_agent
        self._tls = threading.local()

    # ---- 会话（每线程一个，requests.Session 非线程安全）-----------------
    def _build_session(self):
        if not DSKM_HAS_REQUESTS:
            raise RuntimeError("缺少依赖 requests，请先执行：pip install requests")
        session = requests.Session()
        adapter = HTTPAdapter(pool_connections=16, pool_maxsize=32, max_retries=0)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update({
            "Accept": "application/json",
            "User-Agent": self.user_agent,
        })
        return session

    def _session(self):
        session = getattr(self._tls, "session", None)
        if session is None:
            session = self._build_session()
            self._tls.session = session
        session.verify = self.verify_ssl
        if self.proxy_url:
            session.proxies = {"http": self.proxy_url, "https": self.proxy_url}
        else:
            session.proxies = {}
        return session

    def close(self) -> None:
        session = getattr(self._tls, "session", None)
        if session is not None:
            try:
                session.close()
            except Exception:
                pass
            self._tls.session = None

    @property
    def balance_url(self) -> str:
        return f"{self.api_base}{DSKM_BALANCE_PATH}"

    # ---- 错误映射 -------------------------------------------------------
    @staticmethod
    def describe_error(status: int, body_text: str = "") -> str:
        hint = DSKM_HTTP_HINTS.get(int(status or 0), "")
        detail = ""
        text = (body_text or "").strip()
        if text:
            parsed = None
            try:
                parsed = json.loads(text)
            except Exception:
                parsed = None
            if isinstance(parsed, dict):
                err = parsed.get("error")
                if isinstance(err, dict):
                    detail = str(err.get("message") or err.get("type") or "")
                else:
                    detail = str(parsed.get("message") or parsed.get("error") or "")
            if not detail:
                detail = dskm_elide(text, 180)
        if hint and detail:
            return f"{hint}：{detail}"
        if hint:
            return hint
        if detail:
            return f"HTTP {status}：{detail}"
        return f"HTTP {status}"

    @staticmethod
    def classify_error(status: int) -> str:
        if status in (401, 403):
            return "auth"
        if status == 429:
            return "rate_limit"
        if 400 <= status < 500:
            return "client"
        if status >= 500:
            return "server"
        return "http"

    # ---- 解析 -----------------------------------------------------------
    @staticmethod
    def parse_balance_payload(payload: Any) -> Tuple[Optional[bool], List[DSKMBalanceInfo]]:
        if not isinstance(payload, dict):
            return None, []
        available = payload.get(DSKM_FIELD_IS_AVAILABLE, None)
        is_available = None if available is None else bool(available)
        infos_raw = payload.get(DSKM_FIELD_BALANCE_INFOS, [])
        infos: List[DSKMBalanceInfo] = []
        if isinstance(infos_raw, list):
            for item in infos_raw:
                if isinstance(item, dict):
                    infos.append(DSKMBalanceInfo.from_api(item))
        return is_available, infos

    # ---- 主查询 ---------------------------------------------------------
    def fetch_balance(
        self,
        api_key: str,
        uid: str = "",
        cancel_event: Optional[threading.Event] = None,
        progress_cb: Optional[Callable[[int], None]] = None,
    ) -> DSKMQueryResult:
        started = time.time()
        result = DSKMQueryResult(
            uid=uid or dskm_key_uid(api_key),
            checked_at=dskm_now_text(),
        )

        def _emit(pct: int) -> None:
            if progress_cb is not None:
                try:
                    progress_cb(max(0, min(100, int(pct))))
                except Exception:
                    pass

        def _finish() -> DSKMQueryResult:
            result.elapsed_ms = int((time.time() - started) * 1000)
            if not result.checked_at:
                result.checked_at = dskm_now_text()
            return result

        key = dskm_normalize_key(api_key)
        if not key:
            result.error_message = "API Key 为空"
            result.error_kind = "input"
            return _finish()

        if cancel_event is not None and cancel_event.is_set():
            result.error_message = "已取消"
            result.error_kind = "cancelled"
            return _finish()

        if not DSKM_HAS_REQUESTS:
            result.error_message = "缺少依赖 requests，请执行：pip install requests"
            result.error_kind = "dependency"
            return _finish()

        _emit(10)
        attempts = self.max_retry + 1
        last_error = ""
        last_kind = ""
        last_status = 0

        for attempt in range(attempts):
            if cancel_event is not None and cancel_event.is_set():
                result.error_message = "已取消"
                result.error_kind = "cancelled"
                return _finish()
            if attempt > 0:
                _emit(20 + min(30, attempt * 10))
                # 指数退避（1s、2s、4s…上限 8s），期间可被取消打断
                delay = min(8.0, 1.0 * (2 ** (attempt - 1)))
                slept = 0.0
                while slept < delay:
                    if cancel_event is not None and cancel_event.is_set():
                        result.error_message = "已取消"
                        result.error_kind = "cancelled"
                        return _finish()
                    time.sleep(0.2)
                    slept += 0.2
            try:
                _emit(min(65, 30 + attempt * 10))
                session = self._session()
                response = session.get(
                    self.balance_url,
                    headers={"Authorization": f"Bearer {key}"},
                    timeout=(min(10, self.timeout), self.timeout),
                )
                last_status = int(response.status_code)
                result.http_status = last_status
                _emit(80)
                if last_status == 200:
                    try:
                        payload = response.json()
                    except Exception:
                        result.error_message = "接口返回内容不是合法 JSON，可能是网络中间层篡改"
                        result.error_kind = "parse"
                        result.raw = {"text": dskm_elide(response.text, 300)}
                        return _finish()
                    result.raw = payload if isinstance(payload, dict) else {"data": payload}
                    is_available, infos = self.parse_balance_payload(payload)
                    result.ok = True
                    result.is_available = is_available
                    result.balance_infos = infos
                    result.error_message = ""
                    result.error_kind = ""
                    _emit(100)
                    return _finish()

                body = ""
                try:
                    body = response.text or ""
                except Exception:
                    body = ""
                last_error = self.describe_error(last_status, body)
                last_kind = self.classify_error(last_status)
                # 401/403/402 属于确定性错误，无需重试
                if last_status in (401, 402, 403, 404, 405, 422):
                    break
                if last_status < 500 and last_status != 429:
                    break
            except Exception as exc:  # 网络异常
                name = type(exc).__name__
                if DSKM_HAS_REQUESTS and isinstance(exc, requests.exceptions.ConnectTimeout):
                    last_error = f"连接超时（>{self.timeout}s），请检查网络或代理"
                    last_kind = "timeout"
                elif DSKM_HAS_REQUESTS and isinstance(exc, requests.exceptions.ReadTimeout):
                    last_error = f"读取超时（>{self.timeout}s），可尝试增大超时时间"
                    last_kind = "timeout"
                elif DSKM_HAS_REQUESTS and isinstance(exc, requests.exceptions.ProxyError):
                    last_error = "代理连接失败，请检查代理地址与端口"
                    last_kind = "proxy"
                elif DSKM_HAS_REQUESTS and isinstance(exc, requests.exceptions.SSLError):
                    last_error = "SSL 证书校验失败，可尝试关闭“校验 SSL 证书”"
                    last_kind = "ssl"
                elif DSKM_HAS_REQUESTS and isinstance(exc, requests.exceptions.ConnectionError):
                    last_error = f"无法连接服务器（{exc.__class__.__name__}），请检查网络/DNS/防火墙"
                    last_kind = "network"
                else:
                    last_error = f"{name}: {exc}"
                    last_kind = "exception"
                # 取消类异常不重试
                if cancel_event is not None and cancel_event.is_set():
                    result.error_message = "已取消"
                    result.error_kind = "cancelled"
                    return _finish()

        result.error_message = last_error or "查询失败"
        result.error_kind = last_kind or "unknown"
        _emit(100)
        return _finish()


# ============================================================================
#  区域 9：多线程查询（QRunnable + QThreadPool）
# ============================================================================


class DSKMQuerySignals(QObject):
    """QRunnable 不能定义信号，用独立 QObject 承载。"""

    job_started = Signal(str)                 # uid
    job_progress = Signal(str, int)           # uid, percent
    job_finished = Signal(str, object)        # uid, DSKMQueryResult


class DSKMKeyQueryWorker(QRunnable):
    """单密钥查询任务。"""

    def __init__(
        self,
        api_key: str,
        uid: str,
        client: DSKMDeepSeekClient,
        signals: DSKMQuerySignals,
        cancel_event: threading.Event,
    ) -> None:
        super().__init__()
        self.api_key = api_key
        self.uid = uid
        self.client = client
        self.signals = signals
        self.cancel_event = cancel_event
        self.setAutoDelete(True)

    @Slot()
    def run(self) -> None:  # noqa: D102
        try:
            self.signals.job_started.emit(self.uid)
        except Exception:
            pass
        try:
            result = self.client.fetch_balance(
                self.api_key,
                uid=self.uid,
                cancel_event=self.cancel_event,
                progress_cb=lambda pct: self._emit_progress(pct),
            )
        except Exception as exc:  # 兜底：worker 内任何异常都不能吞掉进度
            result = DSKMQueryResult(
                uid=self.uid,
                ok=False,
                error_message=f"内部错误：{type(exc).__name__}: {exc}",
                error_kind="internal",
                checked_at=dskm_now_text(),
            )
        finally:
            try:
                self.client.close()
            except Exception:
                pass
        try:
            self.signals.job_finished.emit(self.uid, result)
        except Exception:
            pass

    def _emit_progress(self, pct: int) -> None:
        try:
            self.signals.job_progress.emit(self.uid, int(pct))
        except Exception:
            pass


class DSKMQueryController(QObject):
    """批量查询调度器：维护线程池、并发数、总进度与取消标志。"""

    batch_started = Signal(int)               # total
    batch_progress = Signal(int, int)         # done, total
    batch_finished = Signal(int, int, float)  # ok, fail, elapsed_seconds
    job_started = Signal(str)
    job_progress = Signal(str, int)
    job_finished = Signal(str, object)

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self.signals = DSKMQuerySignals(self)
        self.signals.job_started.connect(self._relay_started)
        self.signals.job_progress.connect(self._relay_progress)
        self.signals.job_finished.connect(self._relay_finished)
        self._pool = QThreadPool(self)
        self._pool.setMaxThreadCount(5)
        self._cancel_event = threading.Event()
        self._pending = 0
        self._total = 0
        self._done = 0
        self._ok = 0
        self._fail = 0
        self._started_at = 0.0
        self._running = False

    # ---- 配置 -----------------------------------------------------------
    def set_concurrency(self, value: int) -> None:
        self._pool.setMaxThreadCount(max(1, min(64, int(value or 1))))

    def concurrency(self) -> int:
        return self._pool.maxThreadCount()

    def is_running(self) -> bool:
        return self._running or self._pending > 0

    def pending_count(self) -> int:
        return self._pending

    def progress_snapshot(self) -> Tuple[int, int]:
        return self._done, self._total

    # ---- 启动 / 取消 ----------------------------------------------------
    def start_batch(self, jobs: Sequence[Tuple[str, str]], client: DSKMDeepSeekClient) -> int:
        """jobs = [(uid, api_key), ...]，返回实际入队数量。"""
        self.cancel()
        self._cancel_event = threading.Event()
        self._pending = 0
        self._done = 0
        self._ok = 0
        self._fail = 0
        self._total = len(jobs)
        self._started_at = time.time()
        if not jobs:
            self._running = False
            return 0
        self._running = True
        self.batch_started.emit(self._total)
        for uid, api_key in jobs:
            worker = DSKMKeyQueryWorker(api_key, uid, client, self.signals, self._cancel_event)
            self._pending += 1
            self._pool.start(worker)
        return len(jobs)

    def cancel(self) -> None:
        try:
            self._cancel_event.set()
        except Exception:
            pass
        self._pending = 0
        self._running = False

    def wait_all(self, timeout_ms: int = 8000) -> bool:
        return self._pool.waitForDone(max(0, int(timeout_ms)))

    # ---- 内部转发 -------------------------------------------------------
    def _relay_started(self, uid: str) -> None:
        self.job_started.emit(uid)

    def _relay_progress(self, uid: str, pct: int) -> None:
        self.job_progress.emit(uid, pct)

    def _relay_finished(self, uid: str, result: object) -> None:
        self._pending = max(0, self._pending - 1)
        self._done += 1
        if getattr(result, "ok", False):
            self._ok += 1
        elif getattr(result, "error_kind", "") != "cancelled":
            self._fail += 1
        self.job_finished.emit(uid, result)
        self.batch_progress.emit(self._done, self._total)
        if self._pending <= 0:
            self._running = False
            self.batch_finished.emit(self._ok, self._fail, time.time() - self._started_at)


# ============================================================================
#  区域 10：表格列定义 / 表模型 / 过滤代理 / 进度条委托
# ============================================================================

# (key, 表头, 提示, 默认宽度, 对齐, 排序类型)
DSKM_COLUMN_SPECS: List[Tuple[str, str, str, int, str, str]] = [
    ("index", "序号", "当前排序下的行号", 56, "center", "number"),
    ("alias", "名称 / 备注", "为该密钥设置的自定义名称与备注", 170, "left", "text"),
    ("api_key", "Deepseek API KEY", "API Key（默认掩码显示，可在设置中切换）", 250, "left", "text"),
    ("is_available", "账户是否可用\nis_available", "接口字段 is_available：账号是否可用", 118, "center", "bool"),
    ("balance_infos", "余额明细\nbalance_infos", "接口字段 balance_infos：各币种余额明细", 300, "left", "text"),
    ("total_balance", "总可用余额\ntotal_balance", "接口字段 total_balance：总可用余额", 118, "right", "number"),
    ("granted_balance", "赠金余额\ngranted_balance", "接口字段 granted_balance：赠送余额", 118, "right", "number"),
    ("topped_up_balance", "充值余额\ntopped_up_balance", "接口字段 topped_up_balance：充值余额", 128, "right", "number"),
    ("currency", "货币单位\ncurrency", "接口字段 currency：币种", 100, "center", "text"),
    ("status", "查询状态 / 进度", "当前查询进度与结果", 170, "left", "text"),
    ("last_checked", "最后检查时间", "最近一次查询的时间", 160, "center", "text"),
    ("elapsed_ms", "耗时(ms)", "最近一次查询耗时（毫秒）", 90, "right", "number"),
    ("note", "备注", "自定义备注", 160, "left", "text"),
    ("error", "错误 / 说明", "查询失败原因或附加说明", 260, "left", "text"),
]

DSKM_COLUMN_KEYS: List[str] = [spec[0] for spec in DSKM_COLUMN_SPECS]
DSKM_COLUMN_TITLES: Dict[str, str] = {spec[0]: spec[1].replace("\n", " ") for spec in DSKM_COLUMN_SPECS}
DSKM_COLUMN_INDEX: Dict[str, int] = {key: idx for idx, key in enumerate(DSKM_COLUMN_KEYS)}

ROLE_UID = Qt.UserRole + 1
ROLE_PROGRESS = Qt.UserRole + 2
ROLE_STATE = Qt.UserRole + 3
ROLE_SORT = Qt.UserRole + 4


class DSKMKeyTableModel(QAbstractTableModel):
    """密钥表格模型：负责数据缓存、三态排序、按列取值。"""

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._records: List[DSKMKeyRecord] = []
        self._view: List[DSKMKeyRecord] = []
        self._view_pos: Dict[str, int] = {}
        self._insert_pos: Dict[str, int] = {}
        self._sort_key: str = ""
        self._sort_order: str = ""  # 'asc' / 'desc' / '' (默认不排序)
        self._precision: int = 2
        self._key_display_mode: str = KEY_DISPLAY_HEAD_TAIL
        self._key_display_head: int = 6
        self._key_display_tail: int = 4
        self._key_mask_char: str = "*"
        self._key_mask_length: int = 8

    def set_key_display(
        self,
        mode: str = KEY_DISPLAY_HEAD_TAIL,
        head: int = 6,
        tail: int = 4,
        mask_char: str = "*",
        mask_length: int = 8,
    ) -> None:
        """控制 API Key 列的显示方式（完整 / 前N后M / 只前N / 只后M / 全部掩码）。"""
        self._key_display_mode = str(mode or KEY_DISPLAY_HEAD_TAIL)
        self._key_display_head = max(0, min(64, int(head or 0)))
        self._key_display_tail = max(0, min(64, int(tail or 0)))
        self._key_mask_char = (str(mask_char or "*").strip() or "*")[:1]
        try:
            self._key_mask_length = max(0, min(64, int(mask_length or 0)))
        except Exception:
            self._key_mask_length = 8
        self.refresh_all()

    def key_display_state(self) -> Dict[str, Any]:
        return {
            "mode": self._key_display_mode,
            "head": self._key_display_head,
            "tail": self._key_display_tail,
            "mask_char": self._key_mask_char,
            "mask_length": self._key_mask_length,
        }

    def format_key_text(self, api_key: str) -> str:
        """按当前显示设置把 Key 转成可展示文本（列表 / 剪贴板 / 导出 / 日志 均用它）。"""
        return dskm_format_key_display(
            api_key,
            self._key_display_mode,
            self._key_display_head,
            self._key_display_tail,
            self._key_mask_char,
            self._key_mask_length,
        )

    def is_key_hidden(self) -> bool:
        return self._key_display_mode != KEY_DISPLAY_FULL

    def _rebuild_position_maps(self) -> None:
        self._view_pos = {record.uid: idx for idx, record in enumerate(self._view)}
        self._insert_pos = {record.uid: idx for idx, record in enumerate(self._records)}

    # ---- 基础接口 -------------------------------------------------------
    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self._view)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(DSKM_COLUMN_SPECS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):  # noqa: N802
        if orientation != Qt.Horizontal:
            if role == Qt.DisplayRole and 0 <= section < len(self._view):
                return section + 1
            return None
        if role == Qt.DisplayRole and 0 <= section < len(DSKM_COLUMN_SPECS):
            return DSKM_COLUMN_SPECS[section][1]
        if role == Qt.ToolTipRole and 0 <= section < len(DSKM_COLUMN_SPECS):
            return DSKM_COLUMN_SPECS[section][2]
        if role == Qt.TextAlignmentRole and 0 <= section < len(DSKM_COLUMN_SPECS):
            return self._align_flag(DSKM_COLUMN_SPECS[section][4]) | int(Qt.AlignVCenter)
        return None

    @staticmethod
    def _align_flag(align: str) -> int:
        if align == "center":
            return int(Qt.AlignHCenter)
        if align == "right":
            return int(Qt.AlignRight)
        return int(Qt.AlignLeft)

    # ---- 数据 -----------------------------------------------------------
    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):  # noqa: N802
        if not index.isValid():
            return None
        row = index.row()
        col = index.column()
        if not (0 <= row < len(self._view)) or not (0 <= col < len(DSKM_COLUMN_SPECS)):
            return None
        record = self._view[row]
        key = DSKM_COLUMN_SPECS[col][0]
        align = DSKM_COLUMN_SPECS[col][4]

        if role == ROLE_UID:
            return record.uid
        if role == ROLE_PROGRESS:
            if not record.enabled:
                return -1
            if record.last_status == STATUS_RUNNING:
                return max(5, min(99, int(record.progress)))
            if record.last_status in (STATUS_OK, STATUS_FAIL):
                return 100
            return 0
        if role == ROLE_STATE:
            return record.last_status if record.enabled else STATUS_DISABLED
        if role == ROLE_SORT:
            return self.sort_value_of(record, key)

        if role == Qt.TextAlignmentRole:
            return self._align_flag(align) | int(Qt.AlignVCenter)

        if role == Qt.DisplayRole or role == Qt.EditRole:
            return self.display_value_of(record, key)

        if role == Qt.ToolTipRole:
            return self.tooltip_of(record, key)

        if role == Qt.ForegroundRole:
            return self.foreground_of(record, key)

        return None

    # ---- 取值逻辑（表头/导出共用，保证两处一致）-------------------------
    def display_value_of(self, record: DSKMKeyRecord, key: str) -> str:
        if key == "index":
            return str(self._view_pos.get(record.uid, 0) + 1)
        if key == "alias":
            return record.display_alias
        if key == "api_key":
            return self.format_key_text(record.api_key)
        if key == "is_available":
            if record.last_status != STATUS_OK:
                return ""
            return "可用" if record.is_available else "不可用"
        if key == "balance_infos":
            return record.balance_detail_text
        if key == "total_balance":
            return record.total_balance_text
        if key == "granted_balance":
            return record.granted_text
        if key == "topped_up_balance":
            return record.topped_up_text
        if key == "currency":
            return record.currency_text
        if key == "status":
            return self.status_display(record)
        if key == "last_checked":
            return record.last_checked
        if key == "elapsed_ms":
            return str(record.elapsed_ms) if record.elapsed_ms else ""
        if key == "note":
            return record.note
        if key == "error":
            return record.error_message
        return ""

    def status_display(self, record: DSKMKeyRecord) -> str:
        if not record.enabled:
            return STATUS_TEXT[STATUS_DISABLED]
        if record.last_status == STATUS_RUNNING:
            return f"查询中… {max(5, min(99, int(record.progress)))}%"
        if record.last_status == STATUS_OK:
            return "✅ 可用" if record.is_available else "⚠️ 账户不可用"
        if record.last_status == STATUS_FAIL:
            return "❌ 查询失败"
        return "⏳ 待查询"

    def tooltip_of(self, record: DSKMKeyRecord, key: str) -> str:
        base = DSKM_COLUMN_TITLES.get(key, key)
        value = self.display_value_of(record, key)
        lines = [f"{base}：{value or '（空）'}"]
        if key == "api_key":
            lines.append("双击单元格或右键菜单可复制完整 Key")
            if self.is_key_hidden():
                lines.append(
                    "当前显示方式：" + KEY_DISPLAY_MODE_LABELS.get(
                        self._key_display_mode, self._key_display_mode
                    ) + "（可在「设置 → 界面显示 → API Key 显示方式」里改）"
                )
        elif key == "status":
            lines.append(f"状态：{record.status_text}")
            lines.append(f"累计查询 {record.check_count} 次 / 成功 {record.success_count} 次 / 失败 {record.fail_count} 次")
            if record.last_checked:
                lines.append(f"最后检查：{record.last_checked}")
        elif key in ("total_balance", "balance_infos"):
            if record.balance_infos:
                lines.append("明细：" + " ｜ ".join(i.display_text() for i in record.balance_infos))
        elif key == "is_available" and record.is_available is not None:
            lines.append("原始返回值：" + ("true" if record.is_available else "false"))
        return "\n".join(lines)

    def foreground_of(self, record: DSKMKeyRecord, key: str):
        if not record.enabled:
            return QBrush(QColor("#9E9E9E"))
        if record.last_status == STATUS_FAIL:
            if key in ("status", "error", "is_available"):
                return QBrush(QColor("#B71C1C"))
        if record.last_status == STATUS_OK:
            if key == "is_available":
                return QBrush(QColor("#1B5E20") if record.is_available else QBrush(QColor("#B71C1C")))
            if key == "total_balance":
                value = dskm_to_decimal(record.primary_balance.total_balance) if record.primary_balance else None
                if value is not None and value <= 0:
                    return QBrush(QColor("#B71C1C"))
                return QBrush(QColor("#1B5E20"))
        return None

    def sort_value_of(self, record: DSKMKeyRecord, key: str):
        if key == "index":
            return self._insert_pos.get(record.uid, 10 ** 9)
        if key == "alias":
            return record.display_alias
        if key == "api_key":
            return record.api_key
        if key == "is_available":
            if record.is_available is None:
                return None
            return 1 if record.is_available else 0
        if key == "balance_infos":
            return record.balance_detail_text
        if key in ("total_balance", "granted_balance", "topped_up_balance"):
            info = record.primary_balance
            if info is None:
                return None
            value = getattr(info, key)
            dec = dskm_to_decimal(value)
            return float(dec) if dec is not None else None
        if key == "currency":
            return record.currency_text
        if key == "status":
            order = {STATUS_OK: 4, STATUS_RUNNING: 3, STATUS_FAIL: 2, STATUS_PENDING: 1, STATUS_DISABLED: 0}
            return order.get(record.last_status if record.enabled else STATUS_DISABLED, 1)
        if key == "last_checked":
            dt = dskm_parse_time_text(record.last_checked)
            return dt.timestamp() if dt else None
        if key == "elapsed_ms":
            return record.elapsed_ms if record.elapsed_ms else None
        if key == "note":
            return record.note
        if key == "error":
            return record.error_message
        return None

    # ---- 记录管理 -------------------------------------------------------
    def records(self) -> List[DSKMKeyRecord]:
        return list(self._records)

    def view_records(self) -> List[DSKMKeyRecord]:
        return list(self._view)

    def record_at(self, row: int) -> Optional[DSKMKeyRecord]:
        if 0 <= row < len(self._view):
            return self._view[row]
        return None

    def record_by_uid(self, uid: str) -> Optional[DSKMKeyRecord]:
        for record in self._records:
            if record.uid == uid:
                return record
        return None

    def row_of_uid(self, uid: str) -> int:
        for idx, record in enumerate(self._view):
            if record.uid == uid:
                return idx
        return -1

    def set_records(self, records: Sequence[DSKMKeyRecord]) -> None:
        self.beginResetModel()
        self._records = list(records)
        self._apply_view_order(emit=False)
        self.endResetModel()

    def sort_keys(self) -> List[str]:
        return list(DSKM_COLUMN_KEYS)

    def clear_all(self) -> None:
        self.set_records([])

    def append_record(self, record: DSKMKeyRecord) -> None:
        position = len(self._records)
        self.beginInsertRows(QModelIndex(), position, position)
        self._records.append(record)
        self._apply_view_order(emit=False)
        self.endInsertRows()

    def remove_uid(self, uid: str) -> bool:
        for idx, record in enumerate(self._records):
            if record.uid == uid:
                self.beginResetModel()
                self._records.pop(idx)
                self._apply_view_order(emit=False)
                self.endResetModel()
                return True
        return False

    def refresh_uid(self, uid: str) -> None:
        row = self._view_pos.get(uid, -1)
        if row < 0:
            return
        left = self.index(row, 0)
        right = self.index(row, len(DSKM_COLUMN_SPECS) - 1)
        self.dataChanged.emit(left, right)

    def refresh_all(self) -> None:
        if not self._records:
            return
        top = self.index(0, 0)
        bottom = self.index(max(0, len(self._view) - 1), len(DSKM_COLUMN_SPECS) - 1)
        self.dataChanged.emit(top, bottom)

    def update_progress(self, uid: str, status: str, progress: int) -> None:
        record = self.record_by_uid(uid)
        if record is None:
            return
        record.last_status = status
        record.progress = max(0, min(100, int(progress)))
        self.refresh_uid(uid)

    # ---- 排序（三态：升序 → 降序 → 默认不排序）--------------------------
    def sort_state(self) -> Tuple[str, str]:
        return self._sort_key, self._sort_order

    def apply_sort(self, key: str, order: str) -> None:
        self.beginResetModel()
        self._sort_key = key if order else ""
        self._sort_order = order if order else ""
        self._apply_view_order(emit=False)
        self.endResetModel()

    def _apply_view_order(self, emit: bool = True) -> None:
        records = list(self._records)
        order = self._sort_order
        if order and self._sort_key:
            key = self._sort_key
            descending = order == "desc"
            insert_pos = {record.uid: idx for idx, record in enumerate(records)}

            present: List[DSKMKeyRecord] = []
            missing: List[DSKMKeyRecord] = []
            for record in records:
                value = self.sort_value_of(record, key)
                if value is None or (isinstance(value, str) and not value.strip()):
                    missing.append(record)
                else:
                    present.append(record)
            # 数字按数值排（2 < 3 < 10），中文按拼音排；
            # 缺失值永远排在最后（不论升序降序），避免“空值挤在最前”
            present.sort(
                key=lambda r: (_normalize_sort_value(self.sort_value_of(r, key)), insert_pos.get(r.uid, 0)),
                reverse=descending,
            )
            records = present + missing
        self._view = records
        self._rebuild_position_maps()
        if emit:
            self.refresh_all()

    def next_sort_state(self, key: str) -> Tuple[str, str]:
        """点击表头时的状态推进：同列 asc → desc → ''（不排序），换列从 asc 开始。"""
        if self._sort_key == key and self._sort_order == "asc":
            return key, "desc"
        if self._sort_key == key and self._sort_order == "desc":
            return "", ""
        return key, "asc"

    def set_precision(self, precision: int) -> None:
        self._precision = max(0, min(6, int(precision or 0)))
        self.refresh_all()


def _normalize_sort_value(value):
    """排序键规范化：数字按数值排（2 < 3 < 10），中文按拼音排。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value)
    if not text.strip():
        return ""
    number = dskm_to_float(text)
    if number is not None and dskm_is_ascii_digits(text.strip().replace(".", "", 1).lstrip("-")):
        return number
    try:
        from pypinyin import lazy_pinyin
        return "".join(lazy_pinyin(text))
    except Exception:
        return text.lower()


class DSKMKeyFilterProxy(QSortFilterProxyModel):
    """关键词 + 状态 + 币种三重过滤。"""

    FILTER_ALL = "all"

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._search = ""
        self._status = self.FILTER_ALL
        self._currency = self.FILTER_ALL
        self.setDynamicSortFilter(False)

    def set_search_text(self, text: str) -> None:
        value = (text or "").strip().lower()
        if value == self._search:
            return
        self._search = value
        self.invalidateFilter()

    def set_status_filter(self, value: str) -> None:
        value = str(value or self.FILTER_ALL)
        if value == self._status:
            return
        self._status = value
        self.invalidateFilter()

    def set_currency_filter(self, value: str) -> None:
        value = str(value or self.FILTER_ALL)
        if value == self._currency:
            return
        self._currency = value
        self.invalidateFilter()

    def search_text(self) -> str:
        return self._search

    def status_filter(self) -> str:
        return self._status

    def currency_filter(self) -> str:
        return self._currency

    def source_record_at(self, proxy_row: int) -> Optional[DSKMKeyRecord]:
        model = self.sourceModel()
        if not isinstance(model, DSKMKeyTableModel):
            return None
        index = self.mapToSource(self.index(proxy_row, 0))
        if not index.isValid():
            return None
        return model.record_at(index.row())

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:  # noqa: N802
        model = self.sourceModel()
        if not isinstance(model, DSKMKeyTableModel):
            return True
        record = model.record_at(source_row)
        if record is None:
            return False

        # 1) 关键词
        if self._search:
            haystack = " ".join([
                record.api_key.lower(),
                record.masked_key.lower(),
                record.alias.lower(),
                record.note.lower(),
                record.error_message.lower(),
                record.currency_text.lower(),
                record.total_balance_text.lower(),
                record.balance_detail_text.lower(),
            ])
            if self._search not in haystack:
                return False

        # 2) 状态
        status = self._status
        if status != self.FILTER_ALL:
            if status == "enabled" and not record.enabled:
                return False
            elif status == "disabled" and record.enabled:
                return False
            elif status == "available" and not (record.last_status == STATUS_OK and record.is_available):
                return False
            elif status == "unavailable":
                if record.last_status != STATUS_OK or record.is_available:
                    return False
            elif status == "pending" and record.last_status != STATUS_PENDING:
                return False
            elif status == "failed" and record.last_status != STATUS_FAIL:
                return False
            elif status == "success" and record.last_status != STATUS_OK:
                return False
            elif status == "has_balance" and not record.has_balance():
                return False
            elif status == "zero_balance":
                primary = record.primary_balance
                if record.last_status != STATUS_OK or primary is None:
                    return False
                value = dskm_to_decimal(primary.total_balance)
                if value is None or value > 0:
                    return False

        # 3) 币种
        if self._currency != self.FILTER_ALL:
            currencies = [i.currency.upper() for i in record.balance_infos]
            if self._currency.upper() not in currencies:
                return False
        return True


class DSKMProgressDelegate(QStyledItemDelegate):
    """在“查询状态”列绘制一条迷你进度条，直观显示每个密钥的当前进展。"""

    BAR_HEIGHT = 12
    MARGIN = 6

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:  # noqa: D102
        progress = index.data(ROLE_PROGRESS)
        state = index.data(ROLE_STATE)
        if progress is None:
            super().paint(painter, option, index)
            return
        try:
            progress = int(progress)
        except Exception:
            super().paint(painter, option, index)
            return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = option.rect.adjusted(self.MARGIN, 0, -self.MARGIN, 0)
        bar_width = max(40, min(90, rect.width() // 2))
        bar_rect = QRect(
            rect.x() + 2,
            rect.y() + (rect.height() - self.BAR_HEIGHT) // 2,
            bar_width,
            self.BAR_HEIGHT,
        )
        text_rect = QRect(
            bar_rect.right() + 8,
            rect.y(),
            max(10, rect.right() - bar_rect.right() - 8),
            rect.height(),
        )

        # 轨道
        painter.setPen(DSKM_PEN_NO_PEN)
        painter.setBrush(QBrush(QColor("#E3E6EA")))
        painter.drawRoundedRect(bar_rect, 3, 3)

        # 填充
        if progress < 0:
            fill_color = QColor("#BDBDBD")
            fill_width = 0
        elif state == STATUS_FAIL:
            fill_color = QColor("#E57373")
            fill_width = bar_rect.width()
        elif state == STATUS_OK:
            fill_color = QColor("#66BB6A")
            fill_width = bar_rect.width()
        elif state == STATUS_RUNNING:
            fill_color = QColor("#42A5F5")
            fill_width = int(bar_rect.width() * max(0.05, progress / 100.0))
        else:
            fill_color = QColor("#CFD8DC")
            fill_width = int(bar_rect.width() * max(0.0, progress / 100.0))

        if fill_width > 0:
            fill_rect = QRect(bar_rect.x(), bar_rect.y(), fill_width, bar_rect.height())
            painter.setBrush(QBrush(fill_color))
            painter.drawRoundedRect(fill_rect, 3, 3)

        painter.setPen(QPen(QColor("#8A8F98"), 1))
        painter.setBrush(DSKM_BRUSH_NO_BRUSH)
        painter.drawRoundedRect(bar_rect, 3, 3)

        # 文本
        text = str(index.data(Qt.DisplayRole) or "")
        font = option.font
        painter.setFont(font)
        color = option.palette.color(QPalette.Text)
        if state == STATUS_FAIL:
            color = QColor("#B71C1C")
        elif state == STATUS_OK:
            color = QColor("#1B5E20")
        elif state == STATUS_DISABLED:
            color = QColor("#9E9E9E")
        painter.setPen(QPen(color))
        painter.drawText(text_rect, int(Qt.AlignLeft | Qt.AlignVCenter), text)

        try:
            selected = bool(option.state & DSKM_STATE_SELECTED) if DSKM_STATE_SELECTED is not None else False
        except Exception:
            selected = False
        if selected:
            painter.setPen(QPen(option.palette.highlight().color(), 1, DSKM_PEN_DOT_LINE))
            painter.setBrush(DSKM_BRUSH_NO_BRUSH)
            painter.drawRect(option.rect.adjusted(0, 0, -1, -1))

        painter.restore()


# ============================================================================
#  区域 11：通用小控件
# ============================================================================


class DSKMStatCard(QFrame):
    """统计卡片：标题 + 数值 + 副标题。"""

    def __init__(self, title: str, value: str = "-", subtitle: str = "", parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName("dskm_stat_card")
        self.setMinimumHeight(64)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(1)
        self.lbl_title = QLabel(title)
        self.lbl_title.setObjectName("dskm_stat_title")
        self.lbl_value = QLabel(value)
        self.lbl_value.setObjectName("dskm_stat_value")
        self.lbl_sub = QLabel(subtitle)
        self.lbl_sub.setObjectName("dskm_stat_sub")
        self.lbl_sub.setVisible(bool(subtitle))
        layout.addWidget(self.lbl_title)
        layout.addWidget(self.lbl_value)
        layout.addWidget(self.lbl_sub)

    def set_value(self, value: str, subtitle: str = "") -> None:
        self.lbl_value.setText(value)
        self.lbl_sub.setText(subtitle)
        self.lbl_sub.setVisible(bool(subtitle))

    def set_value_color(self, color: str) -> None:
        self.lbl_value.setStyleSheet(f"color: {color};")


class DSKMProgressPanel(QFrame):
    """底部进度面板：状态文本 + 总体进度条 + 计数 + 停止按钮。"""

    cancel_requested = Signal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setMinimumWidth(220)
        self.lbl_status.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")
        self.progress.setMinimumWidth(220)
        self.progress.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self.lbl_count = QLabel("0 / 0")
        self.lbl_count.setMinimumWidth(90)
        self.lbl_count.setAlignment(Qt.AlignCenter)

        self.btn_cancel = QPushButton("停止")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setToolTip("停止当前正在进行的批量查询")
        self.btn_cancel.clicked.connect(self.cancel_requested.emit)

        layout.addWidget(self.lbl_status)
        layout.addWidget(self.progress, 1)
        layout.addWidget(self.lbl_count)
        layout.addWidget(self.btn_cancel)

    def begin(self, total: int, text: str = "") -> None:
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.lbl_count.setText(f"0 / {total}")
        self.lbl_status.setText(text or f"开始查询，共 {total} 个密钥")
        self.btn_cancel.setEnabled(True)

    def update_progress(self, done: int, total: int, text: str = "") -> None:
        total = max(1, int(total))
        self.progress.setValue(int(1000 * min(1.0, done / total)))
        self.lbl_count.setText(f"{done} / {total}")
        if text:
            self.lbl_status.setText(text)

    def finish(self, text: str) -> None:
        self.progress.setValue(1000)
        self.lbl_status.setText(text)
        self.btn_cancel.setEnabled(False)

    def reset(self, text: str = "就绪") -> None:
        self.progress.setValue(0)
        self.lbl_count.setText("0 / 0")
        self.lbl_status.setText(text)
        self.btn_cancel.setEnabled(False)


class DSKMLogPanel(QWidget):
    """日志面板：级别过滤 + 自动滚动 + 清空/复制/导出。"""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        head = QHBoxLayout()
        head.setSpacing(6)
        self.lbl_title = QLabel("运行日志")
        self.cmb_level = QComboBox()
        for key, text in (("debug", "全部"), ("info", "信息及以上"), ("warn", "仅警告与错误"), ("error", "仅错误")):
            self.cmb_level.addItem(text, key)
        self.cmb_level.setCurrentIndex(1)
        self.cmb_level.setFixedWidth(130)
        self.chk_autoscroll = QCheckBox("自动滚动")
        self.chk_autoscroll.setChecked(True)
        self.btn_clear = QPushButton("清空")
        self.btn_copy = QPushButton("复制全部")
        self.btn_open_dir = QPushButton("打开日志目录")
        for btn in (self.btn_clear, self.btn_copy, self.btn_open_dir):
            btn.setMaximumWidth(110)

        head.addWidget(self.lbl_title)
        head.addWidget(QLabel("级别:"))
        head.addWidget(self.cmb_level)
        head.addWidget(self.chk_autoscroll)
        head.addStretch(1)
        head.addWidget(self.btn_copy)
        head.addWidget(self.btn_clear)
        head.addWidget(self.btn_open_dir)

        self.view = QPlainTextEdit()
        self.view.setReadOnly(True)
        self.view.setMaximumBlockCount(3000)
        self.view.setFont(dskm_mono_font(9))
        self.view.setLineWrapMode(QPlainTextEdit.NoWrap)

        layout.addLayout(head)
        layout.addWidget(self.view, 1)

        self.btn_clear.clicked.connect(self.view.clear)
        self.btn_copy.clicked.connect(self._copy_all)
        self.btn_open_dir.clicked.connect(self._open_dir)

    def append_line(self, level: str, text: str) -> None:
        if not self._level_visible(level):
            return
        self.view.appendPlainText(text)
        if self.chk_autoscroll.isChecked():
            bar = self.view.verticalScrollBar()
            bar.setValue(bar.maximum())

    def _level_visible(self, level: str) -> bool:
        mode = self.cmb_level.currentData()
        order = {"debug": 0, "info": 1, "success": 1, "warn": 2, "error": 3}
        if mode == "info":
            return order.get(level, 1) >= 1
        if mode == "warn":
            return order.get(level, 1) >= 2
        if mode == "error":
            return order.get(level, 1) >= 3
        return True

    def _copy_all(self) -> None:
        try:
            QApplication.clipboard().setText(self.view.toPlainText())
            DSKM_LOG.info("日志内容已复制到剪贴板")
        except Exception:
            pass

    def _open_dir(self) -> None:
        try:
            os.startfile(str(dskm_log_dir()))  # noqa: S606 - Windows 打开目录
        except Exception as exc:
            DSKM_LOG.warn(f"打开日志目录失败：{exc}")


# ============================================================================
#  区域 12：对话框
# ============================================================================


class DSKMKeyEditDialog(QDialog):
    """新增 / 编辑单个 API Key（内置连接测试，带进度条）。"""

    def __init__(
        self,
        client_factory: Callable[[], DSKMDeepSeekClient],
        record: Optional[DSKMKeyRecord] = None,
        existing_keys: Optional[Sequence[str]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("新增 API Key" if record is None else "编辑 API Key")
        self.setMinimumWidth(640)
        self._client_factory = client_factory
        self._existing = {dskm_normalize_key(k) for k in (existing_keys or [])}
        self._original_key = dskm_normalize_key(record.api_key) if record else ""
        if self._original_key:
            self._existing.discard(self._original_key)

        self._pool = QThreadPool(self)
        self._pool.setMaxThreadCount(1)
        self._signals = DSKMQuerySignals(self)
        self._signals.job_started.connect(self._on_test_started)
        self._signals.job_progress.connect(self._on_test_progress)
        self._signals.job_finished.connect(self._on_test_finished)
        self._cancel_event = threading.Event()
        self._test_result: Optional[DSKMQueryResult] = None

        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        key_row = QHBoxLayout()
        self.edit_key = QLineEdit()
        self.edit_key.setPlaceholderText("粘贴 sk- 开头的 API Key")
        self.edit_key.setFont(dskm_mono_font(10))
        self.edit_key.setEchoMode(QLineEdit.Password)
        self.btn_show = QToolButton()
        self.btn_show.setText("显示")
        self.btn_show.setCheckable(True)
        self.btn_show.setToolTip("切换明文 / 掩码显示")
        self.btn_show.toggled.connect(self._toggle_echo)
        key_row.addWidget(self.edit_key, 1)
        key_row.addWidget(self.btn_show)
        key_wrap = QWidget()
        key_wrap.setLayout(key_row)
        form.addRow("API Key *：", key_wrap)

        self.edit_alias = QLineEdit()
        self.edit_alias.setPlaceholderText("例如：主账号 / 项目A / 测试专用")
        form.addRow("名称：", self.edit_alias)

        self.edit_note = QLineEdit()
        self.edit_note.setPlaceholderText("任意备注（可选）")
        form.addRow("备注：", self.edit_note)

        self.chk_enabled = QCheckBox("启用该密钥（参与批量查询与自动刷新）")
        self.chk_enabled.setChecked(True)
        form.addRow("状态：", self.chk_enabled)

        layout.addLayout(form)

        # ---- 连接测试区 ----
        test_box = QGroupBox("连接测试")
        test_layout = QVBoxLayout(test_box)
        test_row = QHBoxLayout()
        self.btn_test = QPushButton("测试连接")
        self.btn_test.setToolTip("立即调用官方余额接口验证该 Key")
        self.btn_test.clicked.connect(self._start_test)
        self.progress_test = QProgressBar()
        self.progress_test.setRange(0, 100)
        self.progress_test.setValue(0)
        self.progress_test.setTextVisible(True)
        test_row.addWidget(self.btn_test)
        test_row.addWidget(self.progress_test, 1)
        self.lbl_test = QLabel("尚未测试")
        self.lbl_test.setWordWrap(True)
        self.lbl_test.setTextInteractionFlags(Qt.TextSelectableByMouse)
        test_layout.addLayout(test_row)
        test_layout.addWidget(self.lbl_test)
        layout.addWidget(test_box)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确定")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if record is not None:
            self.edit_key.setText(record.api_key)
            self.edit_alias.setText(record.alias)
            self.edit_note.setText(record.note)
            self.chk_enabled.setChecked(record.enabled)
            if record.last_status == STATUS_OK:
                self.lbl_test.setText(self._describe_record(record))
            elif record.error_message:
                self.lbl_test.setText(f"上次查询失败：{record.error_message}")

    # ---- 交互 -----------------------------------------------------------
    def _toggle_echo(self, shown: bool) -> None:
        self.edit_key.setEchoMode(QLineEdit.Normal if shown else QLineEdit.Password)
        self.btn_show.setText("隐藏" if shown else "显示")

    @staticmethod
    def _describe_record(record: DSKMKeyRecord) -> str:
        if not record.ok_flag:
            return f"上次查询失败：{record.error_message or '未知错误'}"
        flag = "可用" if record.is_available else "不可用"
        detail = record.balance_detail_text or "（接口未返回余额明细）"
        return f"上次查询成功（账户{flag}）｜ {detail}"

    def _start_test(self) -> None:
        key = dskm_normalize_key(self.edit_key.text())
        if not key:
            self.lbl_test.setText("⚠️ 请先填写 API Key")
            return
        if not DSKM_HAS_REQUESTS:
            self.lbl_test.setText("⚠️ 缺少依赖 requests，请执行：pip install requests")
            return
        self._test_result = None
        self.progress_test.setValue(0)
        self.lbl_test.setText("正在测试…")
        self.btn_test.setEnabled(False)
        self._cancel_event = threading.Event()
        try:
            client = self._client_factory()
        except Exception as exc:
            self.lbl_test.setText(f"⚠️ 创建客户端失败：{exc}")
            self.btn_test.setEnabled(True)
            return
        worker = DSKMKeyQueryWorker(key, dskm_key_uid(key), client, self._signals, self._cancel_event)
        self._pool.start(worker)

    def _on_test_started(self, uid: str) -> None:
        self.progress_test.setValue(5)

    def _on_test_progress(self, uid: str, pct: int) -> None:
        self.progress_test.setValue(max(0, min(100, pct)))

    def _on_test_finished(self, uid: str, result: object) -> None:
        self.btn_test.setEnabled(True)
        if not isinstance(result, DSKMQueryResult):
            self.lbl_test.setText("⚠️ 测试异常结束")
            return
        self._test_result = result
        self.progress_test.setValue(100)
        if result.ok:
            flag = "可用" if result.is_available else "不可用（账户被停用）"
            detail = "；".join(i.display_text() for i in result.balance_infos) or "（接口未返回余额明细）"
            self.lbl_test.setText(
                f"✅ 测试成功（{result.elapsed_ms} ms）：账户{flag} ｜ {detail}"
            )
        else:
            self.lbl_test.setText(f"❌ 测试失败（{result.elapsed_ms} ms）：{result.error_message}")

    def _on_accept(self) -> None:
        key = dskm_normalize_key(self.edit_key.text())
        if not key:
            QMessageBox.warning(self, "提示", "请填写 API Key")
            return
        if len(key) < 8:
            QMessageBox.warning(self, "提示", "API Key 长度不足 8 个字符，请检查是否粘贴完整")
            return
        if key in self._existing:
            QMessageBox.warning(self, "提示", "该 API Key 已存在，请勿重复添加")
            return
        self.accept()

    # ---- 取值 -----------------------------------------------------------
    def values(self) -> Dict[str, Any]:
        return {
            "api_key": dskm_normalize_key(self.edit_key.text()),
            "alias": self.edit_alias.text().strip(),
            "note": self.edit_note.text().strip(),
            "enabled": bool(self.chk_enabled.isChecked()),
            "test_result": self._test_result,
        }

    def closeEvent(self, event) -> None:  # noqa: D102, N802
        try:
            self._cancel_event.set()
            self._pool.waitForDone(1500)
        except Exception:
            pass
        super().closeEvent(event)


class DSKMImportDialog(QDialog):
    """导入密钥：支持粘贴文本、选择文件（.env/.txt/.json/.csv）。"""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("导入 API Key")
        self.resize(760, 560)
        self._imported_keys: List[str] = []

        layout = QVBoxLayout(self)
        tip = QLabel(
            "支持：直接粘贴（每行一个 / 逗号分隔）、.env 文件、纯文本、JSON、CSV。\n"
            "程序会自动识别 sk- 开头的 Key 并去重。"
        )
        tip.setWordWrap(True)
        layout.addWidget(tip)

        row = QHBoxLayout()
        self.btn_file = QPushButton("从文件读取…")
        self.btn_env = QPushButton("读取当前 .env")
        self.btn_clip = QPushButton("从剪贴板粘贴")
        self.btn_clear = QPushButton("清空")
        for btn in (self.btn_file, self.btn_env, self.btn_clip, self.btn_clear):
            row.addWidget(btn)
        row.addStretch(1)
        layout.addLayout(row)

        self.edit_source = QPlainTextEdit()
        self.edit_source.setPlaceholderText("在此粘贴 API Key，或点击上方按钮从文件/剪贴板读取…")
        self.edit_source.setFont(dskm_mono_font(9))
        layout.addWidget(self.edit_source, 1)

        info_row = QHBoxLayout()
        self.btn_preview = QPushButton("解析预览")
        self.lbl_info = QLabel("尚未解析")
        info_row.addWidget(self.btn_preview)
        info_row.addWidget(self.lbl_info, 1)
        layout.addLayout(info_row)

        self.list_preview = QListWidget()
        self.list_preview.setFont(dskm_mono_font(9))
        layout.addWidget(self.list_preview, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("导入")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.btn_file.clicked.connect(self._load_from_file)
        self.btn_env.clicked.connect(self._load_from_env)
        self.btn_clip.clicked.connect(self._load_from_clipboard)
        self.btn_clear.clicked.connect(self._clear_source)
        self.btn_preview.clicked.connect(self._refresh_preview)
        self.edit_source.textChanged.connect(self._schedule_preview)
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(350)
        self._preview_timer.timeout.connect(self._refresh_preview)

    def _schedule_preview(self) -> None:
        self._preview_timer.start()

    def _load_from_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择包含 API Key 的文件", "", "所有支持的格式 (*.env *.txt *.json *.csv *.log);;所有文件 (*.*)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as fp:
                self.edit_source.setPlainText(fp.read())
            DSKM_LOG.info(f"已读取导入文件：{path}")
        except Exception as exc:
            QMessageBox.warning(self, "读取失败", f"无法读取文件：{exc}")

    def _load_from_env(self) -> None:
        try:
            store = DSKMEnvStore()
            keys = store.get_api_keys()
            if not keys:
                QMessageBox.information(self, "提示", "当前 .env 中没有可用的 API Key")
                return
            self.edit_source.setPlainText("\n".join(keys))
            DSKM_LOG.info(f"已从 .env 读取 {len(keys)} 个密钥")
        except Exception as exc:
            QMessageBox.warning(self, "读取失败", f"读取 .env 失败：{exc}")

    def _load_from_clipboard(self) -> None:
        try:
            text = QApplication.clipboard().text()
        except Exception:
            text = ""
        if not text.strip():
            QMessageBox.information(self, "提示", "剪贴板为空")
            return
        self.edit_source.setPlainText(text)

    def _clear_source(self) -> None:
        self.edit_source.clear()
        self.list_preview.clear()
        self.lbl_info.setText("尚未解析")

    def _refresh_preview(self) -> None:
        self._imported_keys = dskm_extract_keys(self.edit_source.toPlainText())
        self.list_preview.clear()
        for key in self._imported_keys:
            item = QListWidgetItem(dskm_mask_key(key))
            item.setToolTip("完整 Key 已保存，导入后可查看")
            self.list_preview.addItem(item)
        if self._imported_keys:
            self.lbl_info.setText(f"✅ 解析出 {len(self._imported_keys)} 个唯一密钥")
        else:
            self.lbl_info.setText("⚠️ 未识别到有效密钥")

    def _on_accept(self) -> None:
        self._refresh_preview()
        if not self._imported_keys:
            QMessageBox.warning(self, "提示", "未解析到有效密钥，请检查内容")
            return
        self.accept()

    def imported_keys(self) -> List[str]:
        return list(self._imported_keys)


class DSKMColumnDialog(QDialog):
    """显示列设置：勾选显示哪些列 + 调整顺序。"""

    def __init__(
        self,
        hidden_keys: Sequence[str],
        column_order: Sequence[str],
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("显示列设置")
        self.resize(520, 640)
        self._hidden = set(hidden_keys or [])

        order = [k for k in (column_order or []) if k in DSKM_COLUMN_KEYS]
        for key in DSKM_COLUMN_KEYS:
            if key not in order:
                order.append(key)
        self._order = order

        layout = QVBoxLayout(self)
        tip = QLabel("勾选＝显示该列；取消勾选＝隐藏（仅影响显示，不影响详情/导出）。\n可用「上移 / 下移」调整列顺序。")
        tip.setWordWrap(True)
        layout.addWidget(tip)

        self.list = QListWidget()
        self.list.setMinimumHeight(360)
        layout.addWidget(self.list, 1)

        btn_row = QHBoxLayout()
        self.btn_all = QPushButton("全部显示")
        self.btn_none = QPushButton("全部隐藏")
        self.btn_invert = QPushButton("反选")
        self.btn_reset = QPushButton("恢复默认")
        for btn in (self.btn_all, self.btn_none, self.btn_invert, self.btn_reset):
            btn_row.addWidget(btn)
        layout.addLayout(btn_row)

        move_row = QHBoxLayout()
        self.btn_up = QPushButton("↑ 上移")
        self.btn_down = QPushButton("↓ 下移")
        move_row.addWidget(self.btn_up)
        move_row.addWidget(self.btn_down)
        move_row.addStretch(1)
        layout.addLayout(move_row)

        self.lbl_stats = QLabel("")
        layout.addWidget(self.lbl_stats)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确定")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.btn_all.clicked.connect(lambda: self._set_all(True))
        self.btn_none.clicked.connect(lambda: self._set_all(False))
        self.btn_invert.clicked.connect(self._invert)
        self.btn_reset.clicked.connect(self._reset_default)
        self.btn_up.clicked.connect(lambda: self._move(-1))
        self.btn_down.clicked.connect(lambda: self._move(1))
        self.list.itemChanged.connect(self._update_stats)

        self._rebuild_list()

    def _rebuild_list(self) -> None:
        self.list.blockSignals(True)
        self.list.clear()
        for key in self._order:
            item = QListWidgetItem(DSKM_COLUMN_TITLES.get(key, key))
            item.setData(Qt.UserRole, key)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if key not in self._hidden else Qt.Unchecked)
            self.list.addItem(item)
        self.list.blockSignals(False)
        self._update_stats()

    def _set_all(self, checked: bool) -> None:
        self.list.blockSignals(True)
        for row in range(self.list.count()):
            self.list.item(row).setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.list.blockSignals(False)
        self._update_stats()

    def _invert(self) -> None:
        self.list.blockSignals(True)
        for row in range(self.list.count()):
            item = self.list.item(row)
            item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)
        self.list.blockSignals(False)
        self._update_stats()

    def _reset_default(self) -> None:
        self._hidden = set(DEFAULT_SETTINGS.get(SET_KEY_HIDDEN_COLUMNS, []))
        self._order = list(DSKM_COLUMN_KEYS)
        self._rebuild_list()

    def _move(self, delta: int) -> None:
        row = self.list.currentRow()
        if row < 0:
            return
        target = row + delta
        if not (0 <= target < self.list.count()):
            return
        self._order[row], self._order[target] = self._order[target], self._order[row]
        self._rebuild_list()
        self.list.setCurrentRow(target)

    def _update_stats(self) -> None:
        checked = 0
        for row in range(self.list.count()):
            if self.list.item(row).checkState() == Qt.Checked:
                checked += 1
        self.lbl_stats.setText(f"已勾选 {checked} / {self.list.count()} 列（至少需要显示 1 列）")
        has_index = False
        for row in range(self.list.count()):
            item = self.list.item(row)
            if item.checkState() == Qt.Checked and item.data(Qt.UserRole) == "index":
                has_index = True
        if not has_index and checked > 0:
            pass

    def hidden_columns(self) -> List[str]:
        hidden: List[str] = []
        for row in range(self.list.count()):
            item = self.list.item(row)
            if item.checkState() != Qt.Checked:
                hidden.append(str(item.data(Qt.UserRole)))
        return hidden

    def column_order(self) -> List[str]:
        return [
            str(self.list.item(row).data(Qt.UserRole))
            for row in range(self.list.count())
        ]

    def accept(self) -> None:  # noqa: D102
        if len(self.column_order()) - len(self.hidden_columns()) <= 0:
            QMessageBox.warning(self, "提示", "至少要显示 1 列，否则表格将无法阅读")
            return
        super().accept()


class DSKMEnvEditorDialog(QDialog):
    """直接编辑 .env（API Key 与数据库密码都在这里，不写进代码）。"""

    def __init__(self, env_store: DSKMEnvStore, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("编辑 .env 配置文件")
        self.resize(820, 600)
        self._env = env_store

        layout = QVBoxLayout(self)
        tip = QLabel(
            f"文件位置：<code>{self._env.path}</code><br>"
            "改动保存后立即生效（重新加载密钥即可）。请勿在代码中硬编码密钥。"
        )
        tip.setWordWrap(True)
        layout.addWidget(tip)

        self.edit = QPlainTextEdit()
        self.edit.setFont(dskm_mono_font(10))
        self.edit.setLineWrapMode(QPlainTextEdit.NoWrap)
        layout.addWidget(self.edit, 1)

        row = QHBoxLayout()
        self.btn_reload = QPushButton("重新载入")
        self.btn_mask = QPushButton("隐藏密钥显示")
        self.btn_pick = QPushButton("打开所在目录")
        row.addWidget(self.btn_reload)
        row.addWidget(self.btn_mask)
        row.addWidget(self.btn_pick)
        row.addStretch(1)
        layout.addLayout(row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("保存")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.btn_reload.clicked.connect(self._reload)
        self.btn_mask.clicked.connect(self._toggle_mask)
        self.btn_pick.clicked.connect(self._open_dir)
        self._masked = False
        self._reload()

    def _reload(self) -> None:
        self._env.ensure_exists()
        self.edit.setPlainText(self._env.read_text())

    def _toggle_mask(self) -> None:
        if self._masked:
            self._reload()
            self.btn_mask.setText("隐藏密钥显示")
            self._masked = False
            return
        lines: List[str] = []
        for line in self.edit.toPlainText().splitlines():
            stripped = line.strip()
            if stripped and not stripped.startswith("#") and "=" in stripped:
                key = stripped.split("=", 1)[0].strip()
                if "API_KEY" in key.upper():
                    lines.append(f"{key}={'' if not stripped.split('=', 1)[1].strip() else '••••••••'}")
                    continue
            lines.append(line)
        self.edit.setPlainText("\n".join(lines))
        self.btn_mask.setText("恢复显示")
        self._masked = True

    def _open_dir(self) -> None:
        try:
            os.startfile(str(self._env.path.parent))  # noqa: S606
        except Exception as exc:
            QMessageBox.warning(self, "提示", f"打开目录失败：{exc}")

    def _on_accept(self) -> None:
        text = self.edit.toPlainText()
        if self._masked:
            QMessageBox.warning(self, "提示", "当前为掩码显示，请先点「恢复显示」再保存，避免密钥被覆盖")
            return
        try:
            self._env.write_text(text)
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", f"写入 .env 失败：{exc}")
            return
        DSKM_LOG.success(".env 已保存")
        self.accept()


class DSKMKeyDetailDialog(QDialog):
    """单密钥详情：完整字段、余额明细、历史变化、原始 JSON。"""

    def __init__(
        self,
        record: DSKMKeyRecord,
        precision: int = 2,
        parent: Optional[QWidget] = None,
        key_display: Optional[Dict[str, Any]] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"密钥详情 - {record.display_alias}")
        self.resize(820, 700)
        self._record = record
        self._precision = precision
        self._key_display = dict(key_display or {})
        # 默认按“设置 → API Key 显示方式”隐藏起来；需要时才勾选显示完整 Key
        self._key_mode = str(self._key_display.get("mode", KEY_DISPLAY_HEAD_TAIL))
        self._full_key_text = record.api_key
        self._masked_key_text = dskm_format_key_display(
            record.api_key,
            self._key_mode,
            self._key_display.get("head", 6),
            self._key_display.get("tail", 4),
            self._key_display.get("mask_char", "*"),
            self._key_display.get("mask_length", 8),
        )

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        tabs.setDocumentMode(True)

        # --- 概览 ---
        overview = QWidget()
        grid = QGridLayout(overview)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        self._detail_rows: List[Tuple[str, str]] = [
            ("名称", record.display_alias),
            ("备注", record.note or "（无）"),
            ("API Key", self._masked_key_text),
            ("状态", record.status_text),
            ("账户是否可用 is_available", self._bool_text(record.is_available)),
            ("货币单位 currency", record.currency_text or "（空）"),
            ("总可用余额 total_balance", record.total_balance_text or "（空）"),
            ("赠金余额 granted_balance", record.granted_text or "（空）"),
            ("充值余额 topped_up_balance", record.topped_up_text or "（空）"),
            ("余额明细 balance_infos", record.balance_detail_text or "（空）"),
            ("最后检查时间", record.last_checked or "（从未查询）"),
            ("最近耗时", f"{record.elapsed_ms} ms" if record.elapsed_ms else "（无）"),
            ("累计查询次数", str(record.check_count)),
            ("成功 / 失败", f"{record.success_count} / {record.fail_count}"),
            ("HTTP 状态码", str(record.http_status) if record.http_status else "（无）"),
            ("错误 / 说明", record.error_message or "（无）"),
        ]
        for idx, (label, value) in enumerate(self._detail_rows):
            row = idx
            lbl = QLabel(f"{label}：")
            lbl.setStyleSheet("font-weight: 600;")
            val = QLabel(str(value))
            val.setWordWrap(True)
            val.setTextInteractionFlags(Qt.TextSelectableByMouse)
            if label == "API Key":
                val.setFont(dskm_mono_font(10))
                self.lbl_api_key_value = val
            grid.addWidget(lbl, row, 0, Qt.AlignTop | Qt.AlignRight)
            grid.addWidget(val, row, 1, 1, 3)

        # 默认隐藏，想看清楚再勾（避免截屏/旁人）
        self.chk_reveal_key = QCheckBox("显示完整 Key（默认按「设置 → API Key 显示方式」隐藏）")
        self.chk_reveal_key.setChecked(self._key_mode == KEY_DISPLAY_FULL)
        self.chk_reveal_key.toggled.connect(self._on_reveal_key_toggled)
        grid.addWidget(self.chk_reveal_key, len(self._detail_rows), 1, 1, 3)
        grid.setRowStretch(len(self._detail_rows) + 1, 1)
        tabs.addTab(overview, "概览")

        # --- 余额明细表 ---
        balance_page = QWidget()
        bl = QVBoxLayout(balance_page)
        self.table_balance = QTableWidget(0, 4)
        self.table_balance.setHorizontalHeaderLabels(
            ["货币单位 currency", "总可用余额 total_balance", "赠金余额 granted_balance", "充值余额 topped_up_balance"]
        )
        self.table_balance.verticalHeader().setVisible(False)
        self.table_balance.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_balance.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        if record.balance_infos:
            self.table_balance.setRowCount(len(record.balance_infos))
            for r, info in enumerate(record.balance_infos):
                for c, value in enumerate([
                    info.currency,
                    dskm_format_amount(info.total_balance, precision),
                    dskm_format_amount(info.granted_balance, precision),
                    dskm_format_amount(info.topped_up_balance, precision),
                ]):
                    self.table_balance.setItem(r, c, QTableWidgetItem(str(value)))
        bl.addWidget(self.table_balance, 1)
        tabs.addTab(balance_page, "余额明细 balance_infos")

        # --- 历史 ---
        history_page = QWidget()
        hl = QVBoxLayout(history_page)
        self.table_history = QTableWidget(0, 5)
        self.table_history.setHorizontalHeaderLabels(["时间", "货币", "总余额", "赠金", "充值"])
        self.table_history.verticalHeader().setVisible(False)
        self.table_history.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_history.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        history = list(record.history)
        self.table_history.setRowCount(len(history))
        for r, item in enumerate(reversed(history)):
            values = [
                item.get("time", ""),
                item.get("currency", ""),
                item.get("total", ""),
                item.get("granted", ""),
                item.get("topped_up", ""),
            ]
            for c, value in enumerate(values):
                self.table_history.setItem(r, c, QTableWidgetItem(str(value)))
        hl.addWidget(self.table_history, 1)
        tabs.addTab(history_page, f"余额变化历史（{len(history)}）")

        # --- 原始 JSON ---
        raw_page = QWidget()
        rl = QVBoxLayout(raw_page)
        self.txt_raw = QPlainTextEdit()
        self.txt_raw.setReadOnly(True)
        self.txt_raw.setFont(dskm_mono_font(10))
        self.txt_raw.setPlainText(json.dumps({
            "is_available": record.is_available,
            "balance_infos": [i.to_dict() for i in record.balance_infos],
        }, ensure_ascii=False, indent=2))
        rl.addWidget(self.txt_raw, 1)
        tabs.addTab(raw_page, "接口原始返回")

        layout.addWidget(tabs, 1)

        row = QHBoxLayout()
        self.btn_copy_key = QPushButton("复制完整 Key")
        self.btn_copy_all = QPushButton("复制全部信息")
        self.btn_history = QPushButton("导出历史 CSV")
        row.addWidget(self.btn_copy_key)
        row.addWidget(self.btn_copy_all)
        row.addWidget(self.btn_history)
        row.addStretch(1)
        layout.addLayout(row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("关闭")
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

        self.btn_copy_key.clicked.connect(self._copy_key)
        self.btn_copy_all.clicked.connect(self._copy_all)
        self.btn_history.clicked.connect(self._export_history)

    @staticmethod
    def _bool_text(value: Optional[bool]) -> str:
        if value is None:
            return "（未查询）"
        return "true（可用）" if value else "false（不可用）"

    def _on_reveal_key_toggled(self, checked: bool) -> None:
        """勾上才显示完整 Key，取消则回到掩码。"""
        if hasattr(self, "lbl_api_key_value"):
            self.lbl_api_key_value.setText(
                self._full_key_text if checked else self._masked_key_text
            )

    def _copy_key(self) -> None:
        try:
            QApplication.clipboard().setText(self._record.api_key)
            DSKM_LOG.info("已复制完整 API Key 到剪贴板")
        except Exception:
            pass

    def _copy_all(self) -> None:
        lines = [f"{label}: {value}" for label, value in self._detail_rows]
        try:
            QApplication.clipboard().setText("\n".join(lines))
            DSKM_LOG.info("已复制密钥详情到剪贴板")
        except Exception:
            pass

    def _export_history(self) -> None:
        if not self._record.history:
            QMessageBox.information(self, "提示", "该密钥暂无历史记录")
            return
        default_name = f"密钥历史_{dskm_safe_filename(self._record.display_alias)}_{dskm_now_stamp()}.csv"
        path, _ = QFileDialog.getSaveFileName(self, "导出历史", default_name, "CSV 文件 (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as fp:
                writer = csv.writer(fp)
                writer.writerow(["时间", "货币", "总余额", "赠金", "充值"])
                for item in self._record.history:
                    writer.writerow([
                        item.get("time", ""), item.get("currency", ""),
                        item.get("total", ""), item.get("granted", ""), item.get("topped_up", ""),
                    ])
            DSKM_LOG.success(f"历史记录已导出：{path}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))


class DSKMSettingsDialog(QDialog):
    """设置窗口：所有改动即时生效并自动保存（无需点“确定”）。"""

    setting_changed = Signal(str, object)
    network_setting_changed = Signal(str, object)   # 写入 .env 的网络项
    column_settings_requested = Signal()
    env_editor_requested = Signal()
    reset_requested = Signal()

    def __init__(
        self,
        settings: DSKMSettingsStore,
        env_store: DSKMEnvStore,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.resize(760, 640)
        self._settings = settings
        self._env = env_store
        self._loading = True

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.addTab(self._build_network_tab(), "网络与接口")
        tabs.addTab(self._build_query_tab(), "查询行为")
        tabs.addTab(self._build_ui_tab(), "界面显示")
        tabs.addTab(self._build_data_tab(), "数据与导出")
        layout.addWidget(tabs, 1)

        tip = QLabel("💡 所有设置修改后立即自动保存，下次启动自动恢复。")
        tip.setStyleSheet("color:#2E7D32;")
        layout.addWidget(tip)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("关闭")
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

        self.reload_values_blocking_signals()

    # ---- 构建各页 -------------------------------------------------------
    def _build_network_tab(self) -> QWidget:
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_api_base = QLineEdit()
        self.edit_api_base.setToolTip("默认 https://api.deepseek.com（写入 .env 的 DEEPSEEK_API_BASE）")
        form.addRow("API 基地址：", self.edit_api_base)

        self.spin_timeout = QSpinBox()
        self.spin_timeout.setRange(3, 300)
        self.spin_timeout.setSuffix(" 秒")
        self.spin_timeout.setToolTip("单次请求超时（读取超时），写入 .env 的 DEEPSEEK_TIMEOUT")
        form.addRow("请求超时：", self.spin_timeout)

        self.spin_retry = QSpinBox()
        self.spin_retry.setRange(0, 10)
        self.spin_retry.setSuffix(" 次")
        self.spin_retry.setToolTip("遇到网络错误 / 5xx / 429 时的重试次数（指数退避）")
        form.addRow("失败重试：", self.spin_retry)

        self.spin_concurrency = QSpinBox()
        self.spin_concurrency.setRange(1, 64)
        self.spin_concurrency.setSuffix(" 并发")
        self.spin_concurrency.setToolTip("同时查询的密钥数量；过大可能触发接口限流")
        form.addRow("并发数：", self.spin_concurrency)

        self.chk_verify_ssl = QCheckBox("校验 SSL 证书（遇到自签代理证书问题时才可关闭）")
        form.addRow("安全：", self.chk_verify_ssl)

        self.chk_proxy = QCheckBox("启用 HTTP/HTTPS 代理")
        form.addRow("代理：", self.chk_proxy)

        self.edit_proxy = QLineEdit()
        self.edit_proxy.setPlaceholderText("例如 http://127.0.0.1:7890 或 socks5://127.0.0.1:1080")
        form.addRow("代理地址：", self.edit_proxy)
        self.chk_proxy.toggled.connect(lambda v: self.edit_proxy.setEnabled(bool(v)))

        btn_row = QHBoxLayout()
        self.btn_open_env = QPushButton("编辑 .env 文件…")
        self.btn_test_conn = QPushButton("测试接口连通性（不带 Key）")
        btn_row.addWidget(self.btn_open_env)
        btn_row.addWidget(self.btn_test_conn)
        btn_row.addStretch(1)
        wrap = QWidget()
        wrap.setLayout(btn_row)
        form.addRow("高级：", wrap)

        self.lbl_env_path = QLabel(str(self._env.path))
        self.lbl_env_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(".env 位置：", self.lbl_env_path)

        # 信号
        self.edit_api_base.editingFinished.connect(
            lambda: self._emit_network(SET_KEY_API_BASE, self.edit_api_base.text().strip())
        )
        self.spin_timeout.valueChanged.connect(
            lambda v: self._emit_network(SET_KEY_TIMEOUT, int(v))
        )
        self.spin_retry.valueChanged.connect(
            lambda v: self._emit_network(SET_KEY_MAX_RETRY, int(v))
        )
        self.spin_concurrency.valueChanged.connect(
            lambda v: self._emit_network(SET_KEY_CONCURRENCY, int(v))
        )
        self.chk_verify_ssl.toggled.connect(
            lambda v: self.setting_changed.emit(SET_KEY_VERIFY_SSL, bool(v))
        )
        self.chk_proxy.toggled.connect(
            lambda v: self.setting_changed.emit(SET_KEY_PROXY_ENABLED, bool(v))
        )
        self.edit_proxy.editingFinished.connect(
            lambda: self.setting_changed.emit(SET_KEY_PROXY_URL, self.edit_proxy.text().strip())
        )
        self.btn_open_env.clicked.connect(self.env_editor_requested.emit)
        self.btn_test_conn.clicked.connect(self._test_connectivity)
        return page

    def _build_query_tab(self) -> QWidget:
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)

        self.chk_check_startup = QCheckBox("启动后自动查询一次（延迟 1 秒，避免卡启动）")
        form.addRow("启动查询：", self.chk_check_startup)

        self.chk_auto_refresh = QCheckBox("启用定时自动刷新")
        form.addRow("自动刷新：", self.chk_auto_refresh)

        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(10, 3600)
        self.spin_interval.setSuffix(" 秒")
        form.addRow("刷新间隔：", self.spin_interval)

        self.chk_only_enabled = QCheckBox("仅查询已启用的密钥")
        form.addRow("范围：", self.chk_only_enabled)

        self.chk_skip_recent = QCheckBox("跳过最近刚查过的密钥")
        form.addRow("去重：", self.chk_skip_recent)

        self.spin_skip_seconds = QSpinBox()
        self.spin_skip_seconds.setRange(30, 86400)
        self.spin_skip_seconds.setSuffix(" 秒")
        form.addRow("跳过阈值：", self.spin_skip_seconds)

        self.chk_notify_finish = QCheckBox("批量查询完成后弹窗汇总提示")
        form.addRow("提醒：", self.chk_notify_finish)

        self.chk_keep_history = QCheckBox("记录余额变化历史（详情页可查看趋势）")
        form.addRow("历史：", self.chk_keep_history)

        self.spin_history_limit = QSpinBox()
        self.spin_history_limit.setRange(10, 5000)
        self.spin_history_limit.setSuffix(" 条/密钥")
        form.addRow("历史上限：", self.spin_history_limit)

        self.chk_auto_dedupe = QCheckBox("启动时自动去除重复密钥")
        form.addRow("去重：", self.chk_auto_dedupe)

        for widget, key in (
            (self.chk_check_startup, SET_KEY_CHECK_ON_STARTUP),
            (self.chk_auto_refresh, SET_KEY_AUTO_REFRESH),
            (self.chk_only_enabled, SET_KEY_ONLY_ENABLED_CHECK),
            (self.chk_skip_recent, SET_KEY_SKIP_RECENT_CHECK),
            (self.chk_notify_finish, SET_KEY_NOTIFY_ON_FINISH),
            (self.chk_keep_history, SET_KEY_KEEP_HISTORY),
            (self.chk_auto_dedupe, SET_KEY_AUTO_DEDUPE),
        ):
            widget.toggled.connect(lambda v, k=key: self.setting_changed.emit(k, bool(v)))
        self.spin_interval.valueChanged.connect(
            lambda v: self.setting_changed.emit(SET_KEY_AUTO_REFRESH_INTERVAL, int(v))
        )
        self.spin_skip_seconds.valueChanged.connect(
            lambda v: self.setting_changed.emit(SET_KEY_SKIP_RECENT_SECONDS, int(v))
        )
        self.spin_history_limit.valueChanged.connect(
            lambda v: self.setting_changed.emit(SET_KEY_HISTORY_LIMIT, int(v))
        )
        return page

    def _build_ui_tab(self) -> QWidget:
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)

        self.cmb_theme = QComboBox()
        self.cmb_theme.addItem("浅色（默认）", "light")
        self.cmb_theme.addItem("深色", "dark")
        form.addRow("主题：", self.cmb_theme)

        self.spin_font = QSpinBox()
        self.spin_font.setRange(7, 20)
        self.spin_font.setSuffix(" pt")
        form.addRow("字号：", self.spin_font)

        self.spin_row_height = QSpinBox()
        self.spin_row_height.setRange(20, 80)
        self.spin_row_height.setSuffix(" px")
        form.addRow("行高：", self.spin_row_height)

        self.spin_precision = QSpinBox()
        self.spin_precision.setRange(0, 6)
        self.spin_precision.setSuffix(" 位小数")
        form.addRow("金额精度：", self.spin_precision)

        # ---- API Key 显示方式（★ v1.1.0：列表 / 剪贴板 / 导出 共用一个设置）----
        self.cmb_key_display = QComboBox()
        for mode, label in KEY_DISPLAY_MODE_OPTIONS:
            self.cmb_key_display.addItem(label, mode)
        self.cmb_key_display.setToolTip(
            "列表、剪贴板、导出里的 API Key 都按这里的方式显示，方便隐藏重要信息。"
        )
        form.addRow("API Key 显示：", self.cmb_key_display)

        self.spin_key_head = QSpinBox()
        self.spin_key_head.setRange(0, 32)
        self.spin_key_head.setPrefix("前 ")
        self.spin_key_head.setSuffix(" 位")
        self.spin_key_head.setToolTip("保留开头多少位明文（0 = 一位不留）")
        self.spin_key_tail = QSpinBox()
        self.spin_key_tail.setRange(0, 32)
        self.spin_key_tail.setPrefix("后 ")
        self.spin_key_tail.setSuffix(" 位")
        self.spin_key_tail.setToolTip("保留末尾多少位明文（0 = 一位不留）")
        head_tail_row = QHBoxLayout()
        head_tail_row.setContentsMargins(0, 0, 0, 0)
        head_tail_row.addWidget(self.spin_key_head)
        head_tail_row.addWidget(self.spin_key_tail)
        head_tail_row.addStretch(1)
        head_tail_wrap = QWidget()
        head_tail_wrap.setLayout(head_tail_row)
        form.addRow("保留位数：", head_tail_wrap)

        self.cmb_mask_char = QComboBox()
        for char, label in KEY_MASK_CHAR_OPTIONS:
            self.cmb_mask_char.addItem(label, char)
        self.cmb_mask_char.setToolTip("被隐去的字符用什么符号顶替")
        form.addRow("掩码符号：", self.cmb_mask_char)

        self.spin_mask_length = QSpinBox()
        self.spin_mask_length.setRange(0, 64)
        self.spin_mask_length.setSuffix(" 个")
        self.spin_mask_length.setToolTip(
            "固定个数不会泄露 Key 的真实长度（更安全）；填 0 则按被隐藏的字符数等量显示。"
        )
        form.addRow("掩码长度：", self.spin_mask_length)

        self.lbl_key_preview = QLabel("")
        self.lbl_key_preview.setFont(dskm_mono_font(10))
        self.lbl_key_preview.setWordWrap(True)
        self.lbl_key_preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_key_preview.setToolTip("示例 Key：" + KEY_DISPLAY_SAMPLE)
        form.addRow("效果预览：", self.lbl_key_preview)

        self.chk_stats_bar = QCheckBox("显示顶部统计卡片")
        form.addRow("统计栏：", self.chk_stats_bar)

        self.chk_on_top = QCheckBox("窗口置顶")
        form.addRow("置顶：", self.chk_on_top)

        self.chk_log_visible = QCheckBox("显示日志面板")
        form.addRow("日志：", self.chk_log_visible)

        self.chk_log_autoscroll = QCheckBox("日志自动滚动到末尾")
        form.addRow("", self.chk_log_autoscroll)

        self.spin_log_lines = QSpinBox()
        self.spin_log_lines.setRange(200, 50000)
        self.spin_log_lines.setSuffix(" 行")
        form.addRow("日志上限：", self.spin_log_lines)

        col_row = QHBoxLayout()
        self.btn_columns = QPushButton("显示列设置…")
        self.btn_reset_ui = QPushButton("重置界面布局")
        col_row.addWidget(self.btn_columns)
        col_row.addWidget(self.btn_reset_ui)
        col_row.addStretch(1)
        wrap = QWidget()
        wrap.setLayout(col_row)
        form.addRow("列：", wrap)

        self.cmb_theme.currentIndexChanged.connect(
            lambda: self.setting_changed.emit(SET_KEY_THEME, self.cmb_theme.currentData())
        )
        self.spin_font.valueChanged.connect(lambda v: self.setting_changed.emit(SET_KEY_FONT_SIZE, int(v)))
        self.spin_row_height.valueChanged.connect(lambda v: self.setting_changed.emit(SET_KEY_ROW_HEIGHT, int(v)))
        self.spin_precision.valueChanged.connect(lambda v: self.setting_changed.emit(SET_KEY_CURRENCY_PRECISION, int(v)))
        self.cmb_mask_char.currentIndexChanged.connect(self._on_mask_char_changed)
        self.cmb_key_display.currentIndexChanged.connect(self._on_key_display_changed)
        self.spin_key_head.valueChanged.connect(self._on_key_head_changed)
        self.spin_key_tail.valueChanged.connect(self._on_key_tail_changed)
        self.spin_mask_length.valueChanged.connect(self._on_mask_length_changed)
        self.chk_stats_bar.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_STATISTICS_BAR, bool(v)))
        self.chk_on_top.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_WINDOW_ON_TOP, bool(v)))
        self.chk_log_visible.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_LOG_VISIBLE, bool(v)))
        self.chk_log_autoscroll.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_LOG_AUTOSCROLL, bool(v)))
        self.spin_log_lines.valueChanged.connect(lambda v: self.setting_changed.emit(SET_KEY_LOG_MAX_LINES, int(v)))
        self.btn_columns.clicked.connect(self.column_settings_requested.emit)
        self.btn_reset_ui.clicked.connect(self.reset_requested.emit)
        return page

    def _build_data_tab(self) -> QWidget:
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_export_dir = QLineEdit()
        self.edit_export_dir.setPlaceholderText("留空＝每次导出时询问")
        browse_export = QPushButton("浏览…")
        browse_export.clicked.connect(lambda: self._pick_dir(self.edit_export_dir, SET_KEY_EXPORT_DIR))
        row1 = QHBoxLayout()
        row1.addWidget(self.edit_export_dir, 1)
        row1.addWidget(browse_export)
        wrap1 = QWidget()
        wrap1.setLayout(row1)
        form.addRow("默认导出目录：", wrap1)

        self.edit_import_dir = QLineEdit()
        self.edit_import_dir.setPlaceholderText("留空＝每次导入时询问")
        browse_import = QPushButton("浏览…")
        browse_import.clicked.connect(lambda: self._pick_dir(self.edit_import_dir, SET_KEY_IMPORT_DIR))
        row2 = QHBoxLayout()
        row2.addWidget(self.edit_import_dir, 1)
        row2.addWidget(browse_import)
        wrap2 = QWidget()
        wrap2.setLayout(row2)
        form.addRow("默认导入目录：", wrap2)

        self.chk_export_key = QCheckBox("导出文件包含完整 API Key（默认不包含，更安全）")
        form.addRow("导出内容：", self.chk_export_key)

        self.chk_confirm_delete = QCheckBox("删除密钥前确认")
        form.addRow("确认：", self.chk_confirm_delete)

        self.chk_confirm_clear = QCheckBox("清空全部密钥前确认")
        form.addRow("", self.chk_confirm_clear)

        self.lbl_settings_path = QLabel(str(self._settings.path))
        self.lbl_settings_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow("设置文件：", self.lbl_settings_path)

        self.lbl_meta_path = QLabel(str(dskm_meta_path()))
        self.lbl_meta_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow("密钥元数据：", self.lbl_meta_path)

        row3 = QHBoxLayout()
        self.btn_open_data_dir = QPushButton("打开数据目录")
        self.btn_open_log_dir = QPushButton("打开日志目录")
        row3.addWidget(self.btn_open_data_dir)
        row3.addWidget(self.btn_open_log_dir)
        row3.addStretch(1)
        wrap3 = QWidget()
        wrap3.setLayout(row3)
        form.addRow("目录：", wrap3)

        self.edit_export_dir.editingFinished.connect(
            lambda: self.setting_changed.emit(SET_KEY_EXPORT_DIR, self.edit_export_dir.text().strip())
        )
        self.edit_import_dir.editingFinished.connect(
            lambda: self.setting_changed.emit(SET_KEY_IMPORT_DIR, self.edit_import_dir.text().strip())
        )
        self.chk_export_key.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_EXPORT_INCLUDE_KEY, bool(v)))
        self.chk_confirm_delete.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_CONFIRM_ON_DELETE, bool(v)))
        self.chk_confirm_clear.toggled.connect(lambda v: self.setting_changed.emit(SET_KEY_CONFIRM_ON_CLEAR, bool(v)))
        self.btn_open_data_dir.clicked.connect(lambda: self._open_path(dskm_data_dir()))
        self.btn_open_log_dir.clicked.connect(lambda: self._open_path(dskm_log_dir()))
        return page

    # ---- 辅助 -----------------------------------------------------------
    def _pick_dir(self, target: QLineEdit, key: str) -> None:
        current = target.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "选择目录", current)
        if path:
            target.setText(path)
            self.setting_changed.emit(key, path)

    def _open_path(self, path: Path) -> None:
        try:
            os.startfile(str(path))  # noqa: S606
        except Exception as exc:
            QMessageBox.warning(self, "提示", f"打开目录失败：{exc}")

    def _emit_network(self, key: str, value: Any) -> None:
        if self._loading:
            return
        self.network_setting_changed.emit(key, value)

    # ---- API Key 显示方式：控件 ↔ 设置 ↔ 预览 ---------------------------------
    def key_display_config(self) -> Dict[str, Any]:
        """当前控件上的 API Key 显示配置（主窗口也用它来同步预览/表格）。"""
        return {
            "mode": self.cmb_key_display.currentData() or KEY_DISPLAY_HEAD_TAIL,
            "head": int(self.spin_key_head.value()),
            "tail": int(self.spin_key_tail.value()),
            "mask_char": self.cmb_mask_char.currentData() or "*",
            "mask_length": int(self.spin_mask_length.value()),
        }

    def _emit_setting_guarded(self, key: str, value: Any) -> None:
        if not self._loading:
            self.setting_changed.emit(key, value)

    def _on_key_display_changed(self) -> None:
        self._sync_key_display_preview()
        self._emit_setting_guarded(SET_KEY_KEY_DISPLAY_MODE, self.cmb_key_display.currentData())

    def _on_key_head_changed(self, value: int) -> None:
        self._sync_key_display_preview()
        self._emit_setting_guarded(SET_KEY_KEY_DISPLAY_HEAD, int(value))

    def _on_key_tail_changed(self, value: int) -> None:
        self._sync_key_display_preview()
        self._emit_setting_guarded(SET_KEY_KEY_DISPLAY_TAIL, int(value))

    def _on_mask_char_changed(self) -> None:
        self._sync_key_display_preview()
        self._emit_setting_guarded(SET_KEY_KEY_MASK_CHAR, self.cmb_mask_char.currentData())

    def _on_mask_length_changed(self, value: int) -> None:
        self._sync_key_display_preview()
        self._emit_setting_guarded(SET_KEY_KEY_MASK_LENGTH, int(value))

    def _sync_key_display_preview(self) -> None:
        """按当前选项启用/禁用相关控件，并刷新示例预览。"""
        cfg = self.key_display_config()
        mode = cfg["mode"]
        self.spin_key_head.setEnabled(mode in (KEY_DISPLAY_HEAD_TAIL, KEY_DISPLAY_HEAD_ONLY))
        self.spin_key_tail.setEnabled(mode in (KEY_DISPLAY_HEAD_TAIL, KEY_DISPLAY_TAIL_ONLY))
        masked_mode = mode != KEY_DISPLAY_FULL
        self.cmb_mask_char.setEnabled(masked_mode)
        self.spin_mask_length.setEnabled(masked_mode)
        text = dskm_format_key_display(
            KEY_DISPLAY_SAMPLE, mode, cfg["head"], cfg["tail"],
            cfg["mask_char"], cfg["mask_length"],
        )
        if mode == KEY_DISPLAY_FULL:
            note = "（完整显示，不隐藏任何字符）"
        else:
            note = f"（示例共 {len(KEY_DISPLAY_SAMPLE)} 位，显示长度 {len(text)}）"
        self.lbl_key_preview.setText(f"{text}    {note}")

    def _test_connectivity(self) -> None:
        base = self.edit_api_base.text().strip() or DSKM_DEFAULT_API_BASE
        if not DSKM_HAS_REQUESTS:
            QMessageBox.warning(self, "提示", "缺少依赖 requests")
            return
        try:
            session = requests.Session()
            session.verify = bool(self.chk_verify_ssl.isChecked())
            if self.chk_proxy.isChecked() and self.edit_proxy.text().strip():
                url = self.edit_proxy.text().strip()
                session.proxies = {"http": url, "https": url}
            started = time.time()
            response = session.get(f"{base.rstrip('/')}{DSKM_BALANCE_PATH}", timeout=(5, 10))
            elapsed = int((time.time() - started) * 1000)
            DSKM_LOG.info(f"连通性测试：{base}{DSKM_BALANCE_PATH} → HTTP {response.status_code}（{elapsed} ms）")
            if response.status_code in (200, 401, 403):
                QMessageBox.information(
                    self, "连通性正常",
                    f"接口可达：HTTP {response.status_code}（{elapsed} ms）\n"
                    "返回 401 属正常现象，说明未携带有效 Key。",
                )
            else:
                QMessageBox.warning(
                    self, "连通性异常",
                    f"HTTP {response.status_code}（{elapsed} ms）\n{DSKMDeepSeekClient.describe_error(response.status_code, response.text)}",
                )
        except Exception as exc:
            DSKM_LOG.error(f"连通性测试失败：{exc}")
            QMessageBox.critical(self, "测试失败", f"无法连接：{type(exc).__name__}: {exc}")

    def refresh_from_store(self) -> None:
        """外部（重新打开设置窗口时）刷新控件值。"""
        self.reload_values_blocking_signals()

    def reload_values_blocking_signals(self) -> None:
        """载入控件值：期间阻断所有子控件信号，避免“加载即触发保存”把配置写回。"""
        self._loading = True
        try:
            widgets = self.findChildren(QWidget)
        except Exception:
            widgets = []
        for widget in widgets:
            try:
                widget.blockSignals(True)
            except Exception:
                pass
        try:
            self._load_values()
        finally:
            for widget in widgets:
                try:
                    widget.blockSignals(False)
                except Exception:
                    pass
            self._loading = False

    def _load_values(self) -> None:
        s = self._settings
        self.edit_api_base.setText(self._env.get_api_base())
        opts = self._env.get_network_options()
        self.spin_timeout.setValue(int(opts["timeout"]))
        self.spin_retry.setValue(int(opts["max_retry"]))
        self.spin_concurrency.setValue(int(opts["concurrency"]))
        self.chk_verify_ssl.setChecked(s.get_bool(SET_KEY_VERIFY_SSL, True))
        self.chk_proxy.setChecked(s.get_bool(SET_KEY_PROXY_ENABLED, False))
        self.edit_proxy.setText(s.get_str(SET_KEY_PROXY_URL, ""))
        self.edit_proxy.setEnabled(s.get_bool(SET_KEY_PROXY_ENABLED, False))

        self.chk_check_startup.setChecked(s.get_bool(SET_KEY_CHECK_ON_STARTUP, False))
        self.chk_auto_refresh.setChecked(s.get_bool(SET_KEY_AUTO_REFRESH, False))
        self.spin_interval.setValue(s.get_int(SET_KEY_AUTO_REFRESH_INTERVAL, 30))
        self.chk_only_enabled.setChecked(s.get_bool(SET_KEY_ONLY_ENABLED_CHECK, True))
        self.chk_skip_recent.setChecked(s.get_bool(SET_KEY_SKIP_RECENT_CHECK, False))
        self.spin_skip_seconds.setValue(s.get_int(SET_KEY_SKIP_RECENT_SECONDS, 300))
        self.chk_notify_finish.setChecked(s.get_bool(SET_KEY_NOTIFY_ON_FINISH, True))
        self.chk_keep_history.setChecked(s.get_bool(SET_KEY_KEEP_HISTORY, True))
        self.spin_history_limit.setValue(s.get_int(SET_KEY_HISTORY_LIMIT, 200))
        self.chk_auto_dedupe.setChecked(s.get_bool(SET_KEY_AUTO_DEDUPE, True))

        theme = s.get_str(SET_KEY_THEME, "light")
        self.cmb_theme.setCurrentIndex(1 if theme == "dark" else 0)
        self.spin_font.setValue(s.get_int(SET_KEY_FONT_SIZE, 9))
        self.spin_row_height.setValue(s.get_int(SET_KEY_ROW_HEIGHT, 30))
        self.spin_precision.setValue(s.get_int(SET_KEY_CURRENCY_PRECISION, 2))
        mode = s.get_str(SET_KEY_KEY_DISPLAY_MODE, KEY_DISPLAY_HEAD_TAIL)
        mode_index = self.cmb_key_display.findData(mode)
        self.cmb_key_display.setCurrentIndex(mode_index if mode_index >= 0 else 0)
        self.spin_key_head.setValue(s.get_int(SET_KEY_KEY_DISPLAY_HEAD, 6))
        self.spin_key_tail.setValue(s.get_int(SET_KEY_KEY_DISPLAY_TAIL, 4))
        char_index = self.cmb_mask_char.findData(s.get_str(SET_KEY_KEY_MASK_CHAR, "*"))
        self.cmb_mask_char.setCurrentIndex(char_index if char_index >= 0 else 0)
        self.spin_mask_length.setValue(s.get_int(SET_KEY_KEY_MASK_LENGTH, 8))
        self._sync_key_display_preview()
        self.chk_stats_bar.setChecked(s.get_bool(SET_KEY_STATISTICS_BAR, True))
        self.chk_on_top.setChecked(s.get_bool(SET_KEY_WINDOW_ON_TOP, False))
        self.chk_log_visible.setChecked(s.get_bool(SET_KEY_LOG_VISIBLE, True))
        self.chk_log_autoscroll.setChecked(s.get_bool(SET_KEY_LOG_AUTOSCROLL, True))
        self.spin_log_lines.setValue(s.get_int(SET_KEY_LOG_MAX_LINES, 2000))

        self.edit_export_dir.setText(s.get_str(SET_KEY_EXPORT_DIR, ""))
        self.edit_import_dir.setText(s.get_str(SET_KEY_IMPORT_DIR, ""))
        self.chk_export_key.setChecked(s.get_bool(SET_KEY_EXPORT_INCLUDE_KEY, False))
        self.chk_confirm_delete.setChecked(s.get_bool(SET_KEY_CONFIRM_ON_DELETE, True))
        self.chk_confirm_clear.setChecked(s.get_bool(SET_KEY_CONFIRM_ON_CLEAR, True))
        self.lbl_env_path.setText(str(self._env.path))


class DSKMAboutDialog(QDialog):
    """关于对话框。"""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.resize(660, 520)
        layout = QVBoxLayout(self)
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(self._build_html())
        layout.addWidget(browser, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("关闭")
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def _build_html(self) -> str:
        import platform as _platform
        env = dskm_env_path()
        settings = dskm_settings_path()
        meta = dskm_meta_path()
        icon_state = "已加载" if not dskm_icon().isNull() else "未找到（不影响使用）"
        return f"""
        <h2>{DSKM_APP_NAME}</h2>
        <p><b>{__product_en__}</b> &nbsp; 版本 <b>{__version__}</b></p>
        <p>批量管理 DeepSeek API Key 并查询余额信息，支持无上限数量的密钥。
        调用官方接口 <code>/user/balance</code>，完整展示
        <code>is_available</code> / <code>balance_infos</code> / <code>total_balance</code> /
        <code>granted_balance</code> / <code>topped_up_balance</code> / <code>currency</code> 字段。</p>
        <hr>
        <table cellpadding="4">
          <tr><td>运行环境</td><td>Python {_platform.python_version()} / {_platform.system()} {_platform.release()}</td></tr>
          <tr><td>界面框架</td><td>PySide6（Qt for Python）</td></tr>
          <tr><td>程序图标</td><td>{icon_state}</td></tr>
          <tr><td>程序目录</td><td>{dskm_app_dir()}</td></tr>
          <tr><td>数据目录</td><td>{dskm_data_dir()}</td></tr>
          <tr><td>.env 文件</td><td>{env}</td></tr>
          <tr><td>设置文件</td><td>{settings}</td></tr>
          <tr><td>元数据文件</td><td>{meta}</td></tr>
        </table>
        <hr>
        <p style="color:#666;">API Key 与数据库密码全部保存在 .env 中，代码内不含任何密钥。
        表格默认掩码显示密钥，导出时也可选择是否包含完整 Key。</p>
        """


class DSKMHelpDialog(QDialog):
    """帮助文档。"""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("使用帮助")
        self.resize(900, 700)
        layout = QVBoxLayout(self)
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(DSKM_HELP_HTML)
        layout.addWidget(browser, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("关闭")
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


DSKM_HELP_HTML = """
<h2>DeepSeek 密钥余额管理器 —— 使用帮助</h2>

<h3>一、它做什么</h3>
<p>把你手头所有的 DeepSeek API Key 一次性列出来，多线程同时调用官方余额接口，
把每个 Key 的 <b>账户是否可用</b>、<b>余额明细</b>、<b>总可用余额</b>、<b>赠金余额</b>、
<b>充值余额</b>、<b>货币单位</b> 全部摊在表格里。<b>密钥数量没有上限</b>，
你给多少个就查多少个。</p>

<h3>二、Key 存在哪里</h3>
<p>全部存放在程序目录的 <code>.env</code> 文件中（变量名 <code>DEEPSEEK_API_KEYS</code>，
多个 Key 用英文逗号分隔）。<b>代码里没有任何硬编码密钥</b>，数据库密码同样放在 .env。
你可以在「设置 → 网络与接口 → 编辑 .env 文件…」里直接查看和修改。</p>

<h3>三、官方接口与字段</h3>
<table border="1" cellspacing="0" cellpadding="5">
<tr><th>字段</th><th>含义</th><th>表格列</th></tr>
<tr><td><code>is_available</code></td><td>账户当前是否可用（余额是否为 0 不影响该值）</td><td>账户是否可用</td></tr>
<tr><td><code>balance_infos</code></td><td>余额明细数组，每个币种一条</td><td>余额明细</td></tr>
<tr><td><code>total_balance</code></td><td>该币种总可用余额 = 赠金 + 充值</td><td>总可用余额</td></tr>
<tr><td><code>granted_balance</code></td><td>赠送余额</td><td>赠金余额</td></tr>
<tr><td><code>topped_up_balance</code></td><td>充值余额</td><td>充值余额</td></tr>
<tr><td><code>currency</code></td><td>币种（如 CNY、USD）</td><td>货币单位</td></tr>
</table>
<p>请求地址：<code>GET {API_BASE}/user/balance</code>，请求头
<code>Authorization: Bearer &lt;你的 API Key&gt;</code>。该接口<b>不消耗 Token</b>，可放心频繁调用。</p>

<h3>四、常用操作</h3>
<ul>
<li><b>添加密钥</b>：工具栏「添加密钥」或 Ctrl+N；可一次粘贴多个。</li>
<li><b>批量导入</b>：Ctrl+I，支持粘贴、.env、txt、json、csv。</li>
<li><b>开始查询</b>：F5 或工具栏「批量查询」；支持 Ctrl+Shift+F5 强制重查全部。</li>
<li><b>停止查询</b>：Esc 或底部「停止」按钮。已发出的请求会尽快中断。</li>
<li><b>导出结果</b>：Ctrl+E，支持 CSV / Excel / JSON / Markdown 报表。</li>
<li><b>查看详情</b>：双击任意行，或在右键菜单选择「查看详情」。</li>
<li><b>复制 Key</b>：双击 Key 单元格，或右键「复制完整 Key」。</li>
<li><b>隐藏 Key</b>：菜单「视图 → 🔒 API Key 显示方式」可快速切换；Ctrl+H 一键全部换成掩码符号。</li>
<li><b>启用/停用</b>：右键菜单切换；停用的密钥不参与批量查询与自动刷新。</li>
</ul>

<h3>五、快捷键</h3>
<table border="1" cellspacing="0" cellpadding="5">
<tr><td>Ctrl+N</td><td>添加密钥</td><td>Ctrl+I</td><td>批量导入</td></tr>
<tr><td>Ctrl+E</td><td>导出结果</td><td>F5</td><td>查询全部（按设置过滤）</td></tr>
<tr><td>Ctrl+Shift+F5</td><td>强制重查全部</td><td>Esc</td><td>停止查询</td></tr>
<tr><td>Delete</td><td>删除所选密钥</td><td>Ctrl+A</td><td>全选</td></tr>
<tr><td>Ctrl+L</td><td>聚焦搜索框</td><td>Ctrl+,</td><td>打开设置</td></tr>
<tr><td>Ctrl+D</td><td>打开密钥详情</td><td>Ctrl+Shift+C</td><td>复制所选行信息</td></tr>
<tr><td>Ctrl+R</td><td>重新载入 .env 密钥</td><td>F1</td><td>本帮助</td></tr>
<tr><td>Ctrl+H</td><td>隐藏全部 Key（掩码）</td><td>Ctrl+Shift+H</td><td>恢复完整显示</td></tr>
</table>

<h3>六、常见错误与处理</h3>
<ul>
<li><b>401 认证失败</b>：Key 写错、被删除或被禁用。请到
<a href="https://platform.deepseek.com/api_keys">平台密钥页</a>核对。</li>
<li><b>402 余额不足</b>：账户没钱了，充值后可继续调用；余额接口本身仍可访问。</li>
<li><b>429 请求过于频繁</b>：把并发数降到 3 以下，或把重试次数调高。</li>
<li><b>连接超时 / 无法连接</b>：检查网络、DNS、防火墙，或开启代理后重试。</li>
<li><b>SSL 证书校验失败</b>：若使用公司代理，可在设置里临时关闭「校验 SSL 证书」。</li>
<li><b>余额全是 0 但账户可用</b>：说明赠金与充值都用完了，账户状态仍未停用。</li>
</ul>

<h3>七、API Key 显示方式（隐藏重要信息）</h3>
<p>列表、剪贴板、导出、详情默认<b>都不会把完整的 Key 裸露出来</b>。在
「设置 → 界面显示 → API Key 显示方式」里可以选：</p>
<table border="1" cellspacing="0" cellpadding="5">
<tr><th>方式</th><th>效果（示例 Key 为 sk-***已脱敏***）</th></tr>
<tr><td>前 N 位 + 后 M 位（推荐）</td><td><code>sk-123****stuv</code>，N/M 可调（默认 6 / 4）</td></tr>
<tr><td>只显示前 N 位</td><td><code>sk-123****</code></td></tr>
<tr><td>只显示后 M 位</td><td><code>****stuv</code></td></tr>
<tr><td>全部用掩码符号（最安全）</td><td><code>********</code>，一位明文都不留</td></tr>
<tr><td>完整显示（不隐藏）</td><td><code>sk-***已脱敏***</code></td></tr>
</table>
<p>配套选项：</p>
<ul>
<li><b>掩码符号</b>：<code>*</code> / <code>•</code> / <code>●</code> / <code>×</code> / <code>#</code> / <code>·</code>，随便挑。</li>
<li><b>掩码长度</b>：填 0 = 按被隐藏的字符数显示等量符号；填固定个数
（默认 8）= <b>连 Key 有多长都不泄露</b>，更安全。</li>
<li><b>效果预览</b>：设置里直接就有一个示例 Key 的实时预览，改一下就变。</li>
</ul>
<p>快速操作：菜单「视图 → 🔒 API Key 显示方式」可随时切换；<b>Ctrl+H</b> 一键把全部 Key
换成掩码符号，<b>Ctrl+Shift+H</b> 恢复完整显示（再次隐藏会回到你上次用的掩码方式）。</p>
<p>详情对话框里的 API Key 也默认是隐藏的，需要看清楚时勾一下「显示完整 Key」即可。</p>

<h3>八、关于设置</h3>
<p>任何设置改动都会<b>立即自动保存</b>（带 0.5 秒防抖），关闭程序时再强制保存一次，
下次启动自动恢复，包括窗口位置、列宽、显示列、排序、主题、网络参数、API Key 显示方式等等。</p>
"""


# ============================================================================
#  区域 13：主题样式
# ============================================================================


def dskm_build_stylesheet(theme: str) -> str:
    """返回界面样式表（浅色 / 深色）。"""
    if theme == "dark":
        return """
        QWidget { color: #E6E6E6; }
        QMainWindow, QDialog { background: #1E1F22; }
        QLabel { color: #E6E6E6; }
        QFrame#dskm_stat_card {
            background: #2B2D31; border: 1px solid #3A3D42; border-radius: 6px;
        }
        QLabel#dskm_stat_title { color: #9AA0A6; font-size: 11px; }
        QLabel#dskm_stat_value { color: #FFFFFF; font-size: 16px; font-weight: 600; }
        QLabel#dskm_stat_sub { color: #8A9098; font-size: 10px; }
        QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit, QTextBrowser, QListWidget {
            background: #26282C; border: 1px solid #3A3D42; border-radius: 4px;
            padding: 3px 6px; selection-background-color: #3B6EA5; color: #E6E6E6;
        }
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus { border: 1px solid #4C8DFF; }
        QPushButton {
            background: #2F3237; border: 1px solid #45484E; border-radius: 4px;
            padding: 4px 12px; color: #E6E6E6;
        }
        QPushButton:hover { background: #3A3E44; }
        QPushButton:pressed { background: #4C8DFF; color: #FFFFFF; }
        QPushButton:disabled { color: #6E7378; background: #26282C; }
        QTableView {
            background: #232529; alternate-background-color: #272A2E;
            gridline-color: #34373C; border: 1px solid #3A3D42; selection-background-color: #2F4F7F;
            selection-color: #FFFFFF;
        }
        QHeaderView::section {
            background: #303338; color: #D8D8D8; padding: 4px 6px;
            border: none; border-right: 1px solid #3A3D42; border-bottom: 1px solid #3A3D42;
        }
        QTabWidget::pane { border: 1px solid #3A3D42; }
        QTabBar::tab {
            background: #2B2D31; color: #C8C8C8; padding: 6px 14px; border: 1px solid #3A3D42;
            border-bottom: none;
        }
        QTabBar::tab:selected { background: #3B6EA5; color: #FFFFFF; }
        QGroupBox {
            border: 1px solid #3A3D42; border-radius: 5px; margin-top: 10px; padding-top: 8px;
        }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #A8AEB5; }
        QProgressBar {
            border: 1px solid #3A3D42; border-radius: 4px; background: #26282C;
            text-align: center; color: #E6E6E6; height: 18px;
        }
        QProgressBar::chunk { background: #4C8DFF; border-radius: 3px; }
        QMenuBar { background: #232529; color: #E6E6E6; }
        QMenuBar::item:selected { background: #3B6EA5; }
        QMenu { background: #2B2D31; color: #E6E6E6; border: 1px solid #3A3D42; }
        QMenu::item:selected { background: #3B6EA5; }
        QToolBar { background: #232529; border-bottom: 1px solid #3A3D42; spacing: 4px; }
        QStatusBar { background: #232529; color: #B8BEC5; }
        QScrollBar:vertical { background: #232529; width: 12px; }
        QScrollBar::handle:vertical { background: #45484E; border-radius: 6px; min-height: 24px; }
        QScrollBar::add-line, QScrollBar::sub-line { height: 0; width: 0; }
        QScrollBar:horizontal { background: #232529; height: 12px; }
        QScrollBar::handle:horizontal { background: #45484E; border-radius: 6px; min-width: 24px; }
        QCheckBox { color: #E6E6E6; }
        """
    return """
    QMainWindow, QDialog { background: #F4F6F8; }
    QFrame#dskm_stat_card {
        background: #FFFFFF; border: 1px solid #DFE3E8; border-radius: 6px;
    }
    QLabel#dskm_stat_title { color: #6B7280; font-size: 11px; }
    QLabel#dskm_stat_value { color: #111827; font-size: 16px; font-weight: 600; }
    QLabel#dskm_stat_sub { color: #9CA3AF; font-size: 10px; }
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit, QTextBrowser, QListWidget {
        background: #FFFFFF; border: 1px solid #CBD2D9; border-radius: 4px;
        padding: 3px 6px; selection-background-color: #BFDBFE;
    }
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus { border: 1px solid #2563EB; }
    QPushButton {
        background: #FFFFFF; border: 1px solid #CBD2D9; border-radius: 4px;
        padding: 4px 12px; color: #1F2937;
    }
    QPushButton:hover { background: #EFF6FF; border-color: #93C5FD; }
    QPushButton:pressed { background: #2563EB; color: #FFFFFF; }
    QPushButton:disabled { color: #9CA3AF; background: #F3F4F6; }
    QTableView {
        background: #FFFFFF; alternate-background-color: #FAFBFC;
        gridline-color: #E5E7EB; border: 1px solid #DFE3E8;
        selection-background-color: #DBEAFE; selection-color: #111827;
    }
    QHeaderView::section {
        background: #ECEFF1; color: #374151; padding: 4px 6px; font-weight: 600;
        border: none; border-right: 1px solid #DFE3E8; border-bottom: 1px solid #DFE3E8;
    }
    QTabWidget::pane { border: 1px solid #DFE3E8; background: #FFFFFF; }
    QTabBar::tab {
        background: #ECEFF1; color: #4B5563; padding: 6px 14px;
        border: 1px solid #DFE3E8; border-bottom: none;
    }
    QTabBar::tab:selected { background: #FFFFFF; color: #2563EB; font-weight: 600; }
    QGroupBox {
        border: 1px solid #DFE3E8; border-radius: 5px; margin-top: 10px;
        padding-top: 8px; background: #FFFFFF;
    }
    QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #6B7280; }
    QProgressBar {
        border: 1px solid #CBD2D9; border-radius: 4px; background: #F3F4F6;
        text-align: center; color: #374151; height: 18px;
    }
    QProgressBar::chunk { background: #3B82F6; border-radius: 3px; }
    QMenuBar { background: #FFFFFF; }
    QMenuBar::item:selected { background: #DBEAFE; }
    QMenu { background: #FFFFFF; border: 1px solid #DFE3E8; }
    QMenu::item:selected { background: #DBEAFE; }
    QToolBar { background: #FFFFFF; border-bottom: 1px solid #DFE3E8; spacing: 4px; }
    QStatusBar { background: #FFFFFF; color: #4B5563; }
    QScrollBar:vertical { background: #F4F6F8; width: 12px; }
    QScrollBar::handle:vertical { background: #C7CDD4; border-radius: 6px; min-height: 24px; }
    QScrollBar::add-line, QScrollBar::sub-line { height: 0; width: 0; }
    QScrollBar:horizontal { background: #F4F6F8; height: 12px; }
    QScrollBar::handle:horizontal { background: #C7CDD4; border-radius: 6px; min-width: 24px; }
    """


def dskm_apply_dark_palette(app: QApplication, theme: str) -> None:
    """深色主题下同步调色板，保证委托绘制的文字颜色可读。"""
    palette = QPalette()
    if theme == "dark":
        palette.setColor(QPalette.Window, QColor("#1E1F22"))
        palette.setColor(QPalette.WindowText, QColor("#E6E6E6"))
        palette.setColor(QPalette.Base, QColor("#232529"))
        palette.setColor(QPalette.AlternateBase, QColor("#272A2E"))
        palette.setColor(QPalette.Text, QColor("#E6E6E6"))
        palette.setColor(QPalette.Button, QColor("#2F3237"))
        palette.setColor(QPalette.ButtonText, QColor("#E6E6E6"))
        palette.setColor(QPalette.Highlight, QColor("#2F4F7F"))
        palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
        palette.setColor(QPalette.ToolTipBase, QColor("#2B2D31"))
        palette.setColor(QPalette.ToolTipText, QColor("#E6E6E6"))
    else:
        palette.setColor(QPalette.Window, QColor("#F4F6F8"))
        palette.setColor(QPalette.WindowText, QColor("#111827"))
        palette.setColor(QPalette.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.AlternateBase, QColor("#FAFBFC"))
        palette.setColor(QPalette.Text, QColor("#111827"))
        palette.setColor(QPalette.Button, QColor("#FFFFFF"))
        palette.setColor(QPalette.ButtonText, QColor("#1F2937"))
        palette.setColor(QPalette.Highlight, QColor("#DBEAFE"))
        palette.setColor(QPalette.HighlightedText, QColor("#111827"))
    app.setPalette(palette)


# ============================================================================
#  区域 14：主窗口
# ============================================================================


class DSKMMainWindow(QMainWindow):
    """DeepSeek 密钥余额管理器主窗口。"""

    SETTINGS_SAVE_DELAY_MS = 500

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(DSKM_APP_NAME)
        icon = dskm_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setMinimumSize(880, 560)

        # ---- 数据层 ----
        self.settings = DSKMSettingsStore()
        self.env_store = DSKMEnvStore()
        self.env_store.ensure_exists()
        self.meta_store = DSKMKeyMetaStore()

        self.model = DSKMKeyTableModel(self)
        self.proxy = DSKMKeyFilterProxy(self)
        self.proxy.setSourceModel(self.model)

        self.controller = DSKMQueryController(self)
        self._client: Optional[DSKMDeepSeekClient] = None

        # ---- 运行状态 ----
        self._closing = False
        self._loading_ui = True
        self._query_started_at = 0.0
        self._query_targets = 0
        self._last_query_result: Tuple[int, int, float] = (0, 0, 0.0)
        self._pending_meta_save = False
        self._stats_dirty = False

        # ---- 定时器 ----
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(self.SETTINGS_SAVE_DELAY_MS)
        self._save_timer.timeout.connect(self._save_settings_now)

        self._auto_timer = QTimer(self)
        self._auto_timer.timeout.connect(self._on_auto_refresh_tick)

        self._meta_timer = QTimer(self)
        self._meta_timer.setSingleShot(True)
        self._meta_timer.setInterval(1500)
        self._meta_timer.timeout.connect(self._flush_meta_store)

        self._stats_timer = QTimer(self)
        self._stats_timer.setSingleShot(True)
        self._stats_timer.setInterval(400)
        self._stats_timer.timeout.connect(self.update_statistics)

        self._first_shown = False
        self._start_maximized = False

        # ---- 界面 ----
        self.build_ui()
        self.build_menu_bar()
        self.build_toolbar()
        self.build_shortcuts()
        self.wire_signals()

        self._connect_log_broker()
        self.load_settings()
        self._rebuild_client()
        self.reload_keys_from_env(initial=True)

        self._loading_ui = False
        self.update_statistics()
        self.update_action_states()

        DSKM_LOG.info(f"{DSKM_APP_NAME} v{__version__} 启动完成｜数据目录：{dskm_data_dir()}")

    # ==================================================================
    #  界面搭建
    # ==================================================================
    def build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 6, 8, 4)
        root.setSpacing(6)

        # ---- 统计卡片 ----
        self.stats_widget = QWidget()
        stats_layout = QHBoxLayout(self.stats_widget)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setSpacing(8)
        self.card_total = DSKMStatCard("密钥总数")
        self.card_enabled = DSKMStatCard("已启用")
        self.card_available = DSKMStatCard("账户可用")
        self.card_unavailable = DSKMStatCard("不可用 / 失败")
        self.card_pending = DSKMStatCard("待查询")
        self.card_balance = DSKMStatCard("总可用余额合计")
        for card in (self.card_total, self.card_enabled, self.card_available,
                     self.card_unavailable, self.card_pending, self.card_balance):
            stats_layout.addWidget(card, 1)
        root.addWidget(self.stats_widget)

        # ---- 过滤条 ----
        filter_widget = QWidget()
        filter_layout = QHBoxLayout(filter_widget)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(6)

        filter_layout.addWidget(QLabel("搜索："))
        self.edit_search = QLineEdit()
        self.edit_search.setPlaceholderText("按 Key / 名称 / 备注 / 余额 / 错误信息过滤（Ctrl+L 聚焦）")
        self.edit_search.setClearButtonEnabled(True)
        self.edit_search.setMinimumWidth(240)
        filter_layout.addWidget(self.edit_search, 2)

        filter_layout.addWidget(QLabel("状态："))
        self.cmb_status = QComboBox()
        for value, text in (
            ("all", "全部"),
            ("enabled", "已启用"),
            ("disabled", "已停用"),
            ("available", "账户可用"),
            ("unavailable", "账户不可用"),
            ("success", "查询成功"),
            ("failed", "查询失败"),
            ("pending", "待查询"),
            ("has_balance", "有余额"),
            ("zero_balance", "余额为 0"),
        ):
            self.cmb_status.addItem(text, value)
        self.cmb_status.setMinimumWidth(120)
        filter_layout.addWidget(self.cmb_status)

        filter_layout.addWidget(QLabel("币种："))
        self.cmb_currency = QComboBox()
        self.cmb_currency.addItem("全部", "all")
        self.cmb_currency.setMinimumWidth(100)
        filter_layout.addWidget(self.cmb_currency)

        self.lbl_filter_stats = QLabel("")
        self.lbl_filter_stats.setMinimumWidth(140)
        self.lbl_filter_stats.setStyleSheet("color:#6B7280;")
        filter_layout.addWidget(self.lbl_filter_stats)
        root.addWidget(filter_widget)

        # ---- 主体：表格 + 日志 ----
        self.splitter = QSplitter(Qt.Vertical)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)  # 排序由模型三态控制，禁用 Qt 内置排序
        self.table.setWordWrap(False)
        self.table.setShowGrid(True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setCornerButtonEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionsMovable(True)
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)
        header.setDefaultAlignment(Qt.AlignCenter)
        header.setHighlightSections(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setItemDelegateForColumn(
            DSKM_COLUMN_INDEX["status"], DSKMProgressDelegate(self.table)
        )
        self.splitter.addWidget(self.table)

        self.log_panel = DSKMLogPanel()
        self.splitter.addWidget(self.log_panel)
        self.splitter.setStretchFactor(0, 4)
        self.splitter.setStretchFactor(1, 1)
        root.addWidget(self.splitter, 1)

        # ---- 进度面板 ----
        self.progress_panel = DSKMProgressPanel()
        self.progress_panel.cancel_requested.connect(self.stop_query)
        root.addWidget(self.progress_panel)

        # ---- 状态栏 ----
        bar = self.statusBar()
        bar.setSizeGripEnabled(True)
        self.lbl_status_main = QLabel("就绪")
        self.lbl_status_summary = QLabel("无密钥")
        self.lbl_status_summary.setStyleSheet("padding-left:8px;")
        bar.addWidget(self.lbl_status_main, 1)
        bar.addPermanentWidget(self.lbl_status_summary)

    def build_menu_bar(self) -> None:
        menubar = self.menuBar()

        m_file = menubar.addMenu("文件(&F)")
        self._add_action(m_file, "添加密钥…", self.add_key_dialog, "Ctrl+N", "新增一个 API Key 并可选测试连接")
        self._add_action(m_file, "批量导入密钥…", self.import_keys_dialog, "Ctrl+I", "从文本/文件/.env 批量导入")
        self._add_action(m_file, "重新载入 .env 密钥", self.reload_keys_from_env, "Ctrl+R", "从 .env 重新读取全部密钥")
        m_file.addSeparator()
        self._add_action(m_file, "导出结果…", self.export_results_dialog, "Ctrl+E", "导出为 CSV / Excel / JSON / Markdown")
        self._add_action(m_file, "复制所选行信息", self.copy_selected_rows, "Ctrl+Shift+C")
        m_file.addSeparator()
        self._add_action(m_file, "编辑 .env 文件…", self.open_env_editor)
        self._add_action(m_file, "打开程序目录", lambda: self._open_path(dskm_app_dir()))
        m_file.addSeparator()
        self._add_action(m_file, "退出程序", self.close, "Alt+F4")

        m_op = menubar.addMenu("操作(&O)")
        self._add_action(m_op, "批量查询（按设置过滤）", self.start_query, "F5")
        self._add_action(m_op, "强制重查全部", lambda: self.start_query(force=True), "Ctrl+Shift+F5",
                         "忽略“跳过最近查询”缓存，对全部已启用密钥重新查询")
        self._add_action(m_op, "仅查询所选行", lambda: self.start_query(only_selection=True), "Ctrl+F5")
        self._add_action(m_op, "停止查询", self.stop_query, "Esc")
        m_op.addSeparator()
        self._add_action(m_op, "启用所选", lambda: self.set_selected_enabled(True))
        self._add_action(m_op, "停用所选", lambda: self.set_selected_enabled(False))
        self._add_action(m_op, "上移所选", lambda: self.move_selected(-1))
        self._add_action(m_op, "下移所选", lambda: self.move_selected(1))
        m_op.addSeparator()
        self._add_action(m_op, "编辑所选密钥…", self.edit_selected_key, "Ctrl+E")
        self._add_action(m_op, "查看所选详情", self.show_selected_detail, "Ctrl+D")
        self._add_action(m_op, "删除所选密钥", self.delete_selected_keys, "Delete")
        m_op.addSeparator()
        self._add_action(m_op, "清理重复密钥", self.cleanup_duplicate_keys)
        self._add_action(m_op, "清空全部密钥", self.clear_all_keys)

        m_view = menubar.addMenu("视图(&V)")
        self._add_action(m_view, "显示列设置…", self.configure_columns)
        self.act_stats = self._add_action(m_view, "显示统计卡片", self.toggle_statistics_bar, checkable=True)
        self.act_log = self._add_action(m_view, "显示日志面板", self.toggle_log_panel, checkable=True)
        m_view.addSeparator()
        # ---- API Key 显示方式（★ v1.1.0：可快速切换，改完立即生效并自动保存）----
        self._build_key_display_menu(m_view)
        m_view.addSeparator()
        self.act_theme = self._add_action(m_view, "深色主题", self.toggle_theme, checkable=True)
        self.act_top = self._add_action(m_view, "窗口置顶", self.toggle_always_on_top, checkable=True)
        m_view.addSeparator()
        self._add_action(m_view, "刷新表格", self.refresh_table_view, "Ctrl+Shift+R")
        self._add_action(m_view, "重置窗口位置（居中显示）", self.reset_window_position, "Ctrl+Alt+C")

        m_tool = menubar.addMenu("工具(&T)")
        self._add_action(m_tool, "设置…", self.open_settings_dialog, "Ctrl+,")
        self._add_action(m_tool, "显示列设置…", self.configure_columns)
        m_tool.addSeparator()
        self._add_action(m_tool, "打开数据目录", lambda: self._open_path(dskm_data_dir()))
        self._add_action(m_tool, "打开日志目录", lambda: self._open_path(dskm_log_dir()))
        self._add_action(m_tool, "查看日志文件内容…", self.show_log_dialog)
        m_tool.addSeparator()
        self._add_action(m_tool, "DeepSeek 平台 · 密钥管理", lambda: self._open_url(DSKM_CONFIG_URL))
        self._add_action(m_tool, "DeepSeek 平台 · 用量查询", lambda: self._open_url(DSKM_USAGE_URL))
        self._add_action(m_tool, "DeepSeek 平台 · 充值", lambda: self._open_url(DSKM_TOPUP_URL))

        m_help = menubar.addMenu("帮助(&H)")
        self._add_action(m_help, "使用帮助", self.show_help, "F1")
        self._add_action(m_help, "关于", self.show_about)

    def _add_action(
        self,
        menu: QMenu,
        text: str,
        slot: Callable,
        shortcut: str = "",
        tip: str = "",
        checkable: bool = False,
    ) -> QAction:
        action = QAction(text, self)
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        if tip:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        action.setCheckable(checkable)
        action.triggered.connect(slot)
        menu.addAction(action)
        if not hasattr(self, "_all_actions"):
            self._all_actions: List[QAction] = []
        self._all_actions.append(action)
        return action

    def _build_key_display_menu(self, parent_menu: QMenu) -> None:
        """在「视图」菜单里加「API Key 显示方式」子菜单（★ v1.1.0）。

        好处：截图、演示、有人站在身后的时候，一下就能把 Key 全变成 `*`，
        不用钻进设置里翻。改完立即生效并自动保存。
        """
        submenu = parent_menu.addMenu("🔒 API Key 显示方式")
        group = QActionGroup(self)
        group.setExclusive(True)
        for mode, label in KEY_DISPLAY_MODE_OPTIONS:
            action = QAction(label, self)
            action.setCheckable(True)
            action.setData(mode)
            action.setStatusTip(f"列表 / 剪贴板 / 导出里的 API Key 都按「{label}」显示")
            action.triggered.connect(lambda _checked=False, m=mode: self.set_key_display_mode(m))
            group.addAction(action)
            submenu.addAction(action)
            self._all_actions.append(action)
        self._key_display_action_group = group
        submenu.addSeparator()

        self.act_key_hide = QAction("🙈 一键隐藏全部（改为掩码符号）", self)
        self.act_key_hide.setShortcut(QKeySequence("Ctrl+H"))
        self.act_key_hide.setStatusTip("立刻把所有 API Key 变成掩码显示（截图 / 演示前按一下）")
        self.act_key_hide.triggered.connect(lambda: self.toggle_key_visibility(True))
        self.act_key_show = QAction("👁 显示完整 Key（不隐藏）", self)
        self.act_key_show.setShortcut(QKeySequence("Ctrl+Shift+H"))
        self.act_key_show.setStatusTip("恢复完整显示；再次隐藏时会回到你上次用的掩码方式")
        self.act_key_show.triggered.connect(lambda: self.toggle_key_visibility(False))
        for action in (self.act_key_hide, self.act_key_show):
            submenu.addAction(action)
            self._all_actions.append(action)
        submenu.addSeparator()
        self._add_action(submenu, "打开完整设置…", self.open_settings_dialog)

    def build_toolbar(self) -> None:
        toolbar = QToolBar("主工具栏")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(18, 18))
        toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.addToolBar(toolbar)
        self.toolbar = toolbar

        self.act_add = QAction("➕ 添加密钥", self)
        self.act_import = QAction("📥 批量导入", self)
        self.act_query = QAction("🔄 批量查询", self)
        self.act_query_force = QAction("⚡ 强制重查", self)
        self.act_stop = QAction("⏹ 停止", self)
        self.act_export = QAction("📤 导出结果", self)
        self.act_columns = QAction("🗂 显示列", self)
        self.act_settings = QAction("⚙ 设置", self)
        self.act_help = QAction("❓ 帮助", self)
        for action in (self.act_add, self.act_import, self.act_query, self.act_query_force, self.act_stop):
            toolbar.addAction(action)
        toolbar.addSeparator()
        for action in (self.act_export, self.act_columns, self.act_settings):
            toolbar.addAction(action)
        toolbar.addSeparator()
        toolbar.addAction(self.act_help)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        toolbar.addWidget(spacer)
        self.lbl_toolbar_hint = QLabel("  提示：双击行可查看详情｜右键可复制/编辑/删除  ")
        self.lbl_toolbar_hint.setStyleSheet("color:#6B7280;")
        toolbar.addWidget(self.lbl_toolbar_hint)

        self.act_add.triggered.connect(self.add_key_dialog)
        self.act_import.triggered.connect(self.import_keys_dialog)
        self.act_query.triggered.connect(self.start_query)
        self.act_query_force.triggered.connect(lambda: self.start_query(force=True))
        self.act_stop.triggered.connect(self.stop_query)
        self.act_export.triggered.connect(self.export_results_dialog)
        self.act_columns.triggered.connect(self.configure_columns)
        self.act_settings.triggered.connect(self.open_settings_dialog)
        self.act_help.triggered.connect(self.show_help)

    def build_shortcuts(self) -> None:
        # 菜单已带快捷键，这里补充无菜单项的操作
        shortcut = QShortcut(QKeySequence("Ctrl+A"), self.table)
        shortcut.setContext(Qt.WidgetShortcut)
        shortcut.activated.connect(self.table.selectAll)
        shortcut2 = QShortcut(QKeySequence("Ctrl+Shift+R"), self)
        shortcut2.activated.connect(self.refresh_table_view)

    # ==================================================================
    #  信号连接
    # ==================================================================
    def wire_signals(self) -> None:
        self.controller.batch_started.connect(self._on_batch_started)
        self.controller.batch_progress.connect(self._on_batch_progress)
        self.controller.batch_finished.connect(self._on_batch_finished)
        self.controller.job_started.connect(self._on_job_started)
        self.controller.job_progress.connect(self._on_job_progress)
        self.controller.job_finished.connect(self._on_job_finished)

        self.edit_search.textChanged.connect(self._on_search_changed)
        self.cmb_status.currentIndexChanged.connect(self._on_status_filter_changed)
        self.cmb_currency.currentIndexChanged.connect(self._on_currency_filter_changed)

        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)
        self.table.horizontalHeader().sectionResized.connect(self._on_section_resized)
        self.table.horizontalHeader().sectionMoved.connect(lambda *_: self.schedule_settings_save())
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        self.table.customContextMenuRequested.connect(self._on_table_context_menu)
        self.table.selectionModel().selectionChanged.connect(lambda *_: self.update_action_states())

        self.proxy.rowsInserted.connect(lambda *_: self._refresh_filter_stats())
        self.proxy.rowsRemoved.connect(lambda *_: self._refresh_filter_stats())
        self.proxy.layoutChanged.connect(lambda *_: self._refresh_filter_stats())
        self.proxy.modelReset.connect(lambda *_: self._refresh_filter_stats())

    def _connect_log_broker(self) -> None:
        DSKM_LOG.message_logged.connect(self.log_panel.append_line)

    # ==================================================================
    #  键盘/滚轮/双击等交互
    # ==================================================================
    def _on_table_double_clicked(self, index: QModelIndex) -> None:
        if not index.isValid():
            return
        source = self.proxy.mapToSource(index)
        record = self.model.record_at(source.row())
        if record is None:
            return
        column_key = DSKM_COLUMN_KEYS[source.column()] if source.column() < len(DSKM_COLUMN_KEYS) else ""
        if column_key == "api_key":
            self._copy_text(record.api_key, "完整 API Key")
        else:
            self.show_detail_for(record)

    def _on_table_context_menu(self, pos: QPoint) -> None:
        index = self.table.indexAt(pos)
        if index.isValid() and index.row() not in [i.row() for i in self.table.selectionModel().selectedRows()]:
            self.table.selectRow(index.row())
        records = self.selected_records()
        if not records:
            return
        single = records[0] if len(records) == 1 else None

        menu = QMenu(self)
        if single is not None:
            menu.addAction("查看详情…", lambda: self.show_detail_for(single))
            menu.addAction("编辑密钥…", lambda: self.edit_key_record(single))
            menu.addSeparator()
            menu.addAction("复制完整 Key", lambda: self._copy_text(single.api_key, "完整 API Key"))
            menu.addAction("复制按当前方式显示的 Key",
                           lambda: self._copy_text(self.format_key_text(single.api_key), "显示方式处理后的 Key"))
        menu.addAction(f"复制所选行信息（{len(records)} 行）", self.copy_selected_rows)
        menu.addAction("复制表格（TSV）", self.copy_table_as_tsv)
        menu.addSeparator()
        menu.addAction("查询所选", lambda: self.start_query(only_selection=True))
        menu.addSeparator()
        menu.addAction("启用所选", lambda: self.set_selected_enabled(True))
        menu.addAction("停用所选", lambda: self.set_selected_enabled(False))
        menu.addAction("上移", lambda: self.move_selected(-1))
        menu.addAction("下移", lambda: self.move_selected(1))
        menu.addSeparator()
        menu.addAction("删除所选", self.delete_selected_keys)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _on_header_clicked(self, logical_index: int) -> None:
        if not (0 <= logical_index < len(DSKM_COLUMN_KEYS)):
            return
        key = DSKM_COLUMN_KEYS[logical_index]
        new_key, new_order = self.model.next_sort_state(key)
        self.model.apply_sort(new_key, new_order)
        self._update_sort_indicator()
        self.set_setting(SET_KEY_SORT_COLUMN, new_key)
        self.set_setting(SET_KEY_SORT_ORDER, new_order)
        if not new_order:
            self.flash_status(f"已恢复默认顺序（不排序）")
        else:
            title = DSKM_COLUMN_TITLES.get(key, key)
            self.flash_status(f"按「{title}」{'升序' if new_order == 'asc' else '降序'}排序（再点一次可切换/取消）")

    def _update_sort_indicator(self) -> None:
        key, order = self.model.sort_state()
        header = self.table.horizontalHeader()
        if not order or key not in DSKM_COLUMN_INDEX:
            header.setSortIndicatorShown(False)
            return
        header.setSortIndicatorShown(True)
        logical = DSKM_COLUMN_INDEX[key]
        header.setSortIndicator(logical, Qt.AscendingOrder if order == "asc" else Qt.DescendingOrder)

    def _on_section_resized(self, logical_index: int, old_size: int, new_size: int) -> None:
        if self._loading_ui:
            return
        self.schedule_settings_save()

    # ==================================================================
    #  设置读写（实时自动保存）
    # ==================================================================
    def get_setting(self, key: str, default: Any = None) -> Any:
        return self.settings.get(key, default)

    def set_setting(self, key: str, value: Any) -> None:
        """写入设置并立即应用 + 安排落盘（防抖）。"""
        self.settings.set(key, value)
        self.apply_setting(key, value)
        self.schedule_settings_save()

    def schedule_settings_save(self) -> None:
        if self._loading_ui or self._closing:
            return
        self._save_timer.start()

    def _collect_geometry_settings(self) -> Dict[str, Any]:
        payload = {
            SET_KEY_WINDOW_MAXIMIZED: bool(self.isMaximized()),
            SET_KEY_WINDOW_GEOMETRY: dskm_window_geometry_for_save(
                self, self.get_setting(SET_KEY_WINDOW_GEOMETRY, DEFAULT_SETTINGS[SET_KEY_WINDOW_GEOMETRY])
            ),
            SET_KEY_SPLITTER_SIZES: list(self.splitter.sizes()),
        }
        widths: Dict[str, int] = {}
        header = self.table.horizontalHeader()
        for key in DSKM_COLUMN_KEYS:
            logical = DSKM_COLUMN_INDEX.get(key)
            if logical is None:
                continue
            widths[key] = int(header.sectionSize(logical))
        payload[SET_KEY_COLUMN_WIDTHS] = widths
        order: List[str] = []
        for visual in range(len(DSKM_COLUMN_KEYS)):
            logical = header.logicalIndex(visual)
            if 0 <= logical < len(DSKM_COLUMN_KEYS):
                order.append(DSKM_COLUMN_KEYS[logical])
        payload[SET_KEY_COLUMN_ORDER] = order
        payload[SET_KEY_HIDDEN_COLUMNS] = [
            key for key in DSKM_COLUMN_KEYS
            if (DSKM_COLUMN_INDEX.get(key) is not None and self.table.isColumnHidden(DSKM_COLUMN_INDEX[key]))
        ]
        payload[SET_KEY_LOG_VISIBLE] = bool(self.log_panel.isVisible())
        if self.log_panel.isVisible():
            payload[SET_KEY_LOG_HEIGHT] = int(self.log_panel.height())
        payload[SET_KEY_STATISTICS_BAR] = bool(self.stats_widget.isVisible())
        return payload

    def _save_settings_now(self) -> None:
        try:
            self.settings.update(self._collect_geometry_settings())
            ok = self.settings.save()
            if ok:
                DSKM_LOG.debug(f"设置已保存：{self.settings.path}")
        except Exception as exc:
            DSKM_LOG.warn(f"保存设置失败：{exc}")

    def apply_setting(self, key: str, value: Any) -> None:
        """把单个设置应用到界面（新增设置项必须在这里处理）。"""
        try:
            if key == SET_KEY_THEME:
                self.apply_theme(str(value))
            elif key == SET_KEY_FONT_SIZE:
                self.apply_font_size(int(value))
            elif key == SET_KEY_ROW_HEIGHT:
                self.table.verticalHeader().setDefaultSectionSize(int(value))
            elif key == SET_KEY_CURRENCY_PRECISION:
                self.model.set_precision(int(value))
                self.update_statistics()
            elif key in (SET_KEY_KEY_DISPLAY_MODE, SET_KEY_KEY_DISPLAY_HEAD, SET_KEY_KEY_DISPLAY_TAIL,
                         SET_KEY_KEY_MASK_CHAR, SET_KEY_KEY_MASK_LENGTH):
                self.apply_key_display_settings()
            elif key == SET_KEY_STATISTICS_BAR:
                self.stats_widget.setVisible(bool(value))
                if hasattr(self, "act_stats"):
                    self.act_stats.setChecked(bool(value))
            elif key == SET_KEY_LOG_VISIBLE:
                self.log_panel.setVisible(bool(value))
                if hasattr(self, "act_log"):
                    self.act_log.setChecked(bool(value))
            elif key == SET_KEY_LOG_MAX_LINES:
                self.log_panel.view.setMaximumBlockCount(max(100, int(value)))
            elif key == SET_KEY_LOG_AUTOSCROLL:
                self.log_panel.chk_autoscroll.setChecked(bool(value))
            elif key == SET_KEY_WINDOW_ON_TOP:
                self.apply_always_on_top(bool(value))
            elif key == SET_KEY_CONCURRENCY:
                self.controller.set_concurrency(int(value))
            elif key in (SET_KEY_VERIFY_SSL, SET_KEY_PROXY_ENABLED, SET_KEY_PROXY_URL,
                         SET_KEY_API_BASE, SET_KEY_TIMEOUT, SET_KEY_MAX_RETRY):
                self._rebuild_client()
            elif key == SET_KEY_AUTO_REFRESH:
                self._restart_auto_timer()
            elif key == SET_KEY_AUTO_REFRESH_INTERVAL:
                self._restart_auto_timer()
        except Exception as exc:
            DSKM_LOG.warn(f"应用设置 {key} 失败：{exc}")

    # ---- API Key 显示方式（★ v1.1.0）------------------------------------------
    def key_display_config(self) -> Dict[str, Any]:
        """当前密钥显示配置（详情对话框也用它，保证全程序一致）。"""
        return {
            "mode": self.settings.get_str(SET_KEY_KEY_DISPLAY_MODE, KEY_DISPLAY_HEAD_TAIL),
            "head": self.settings.get_int(SET_KEY_KEY_DISPLAY_HEAD, 6),
            "tail": self.settings.get_int(SET_KEY_KEY_DISPLAY_TAIL, 4),
            "mask_char": self.settings.get_str(SET_KEY_KEY_MASK_CHAR, "*"),
            "mask_length": self.settings.get_int(SET_KEY_KEY_MASK_LENGTH, 8),
        }

    def format_key_text(self, api_key: str) -> str:
        return self.model.format_key_text(api_key)

    def apply_key_display_settings(self) -> None:
        """把「API Key 显示方式」应用到表格，并同步菜单里的单选状态。"""
        cfg = self.key_display_config()
        self.model.set_key_display(
            cfg["mode"], cfg["head"], cfg["tail"], cfg["mask_char"], cfg["mask_length"]
        )
        group = getattr(self, "_key_display_action_group", None)
        if group is not None:
            for action in group.actions():
                action.setChecked(action.data() == cfg["mode"])
        if hasattr(self, "act_key_hide") and hasattr(self, "act_key_show"):
            hidden = cfg["mode"] != KEY_DISPLAY_FULL
            self.act_key_hide.setChecked(hidden)
            self.act_key_show.setChecked(not hidden)

    def set_key_display_mode(self, mode: str) -> None:
        """菜单快捷切换显示方式（即时生效 + 自动保存）。"""
        self.set_setting(SET_KEY_KEY_DISPLAY_MODE, mode)
        label = KEY_DISPLAY_MODE_LABELS.get(mode, mode)
        self.flash_status(f"API Key 显示方式已切换为：{label}")
        DSKM_LOG.info(f"API Key 显示方式：{label}")

    def toggle_key_visibility(self, hidden: bool) -> None:
        """一键隐藏 / 显示全部 API Key（隐藏时回到上次用的掩码方式）。"""
        if hidden:
            mode = self.settings.get_str(SET_KEY_LAST_MASKED_MODE, "") or KEY_DISPLAY_ALL_MASK
            if mode == KEY_DISPLAY_FULL:
                mode = KEY_DISPLAY_ALL_MASK
            self.set_key_display_mode(mode)
        else:
            current = self.settings.get_str(SET_KEY_KEY_DISPLAY_MODE, KEY_DISPLAY_HEAD_TAIL)
            if current != KEY_DISPLAY_FULL:
                self.settings.set(SET_KEY_LAST_MASKED_MODE, current)
            self.set_key_display_mode(KEY_DISPLAY_FULL)

    def load_settings(self) -> None:
        """启动时恢复全部设置。"""
        s = self.settings
        # 窗口几何
        geometry = dskm_safe_window_geometry(
            s.get(SET_KEY_WINDOW_GEOMETRY, DEFAULT_SETTINGS[SET_KEY_WINDOW_GEOMETRY]),
            default_size=(1420, 860),
        )
        dskm_place_window(self, geometry)
        self._start_maximized = s.get_bool(SET_KEY_WINDOW_MAXIMIZED, False)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, s.get_bool(SET_KEY_WINDOW_ON_TOP, False))

        # 字体 / 行高 / 精度
        self.apply_font_size(s.get_int(SET_KEY_FONT_SIZE, 9))
        self.table.verticalHeader().setDefaultSectionSize(s.get_int(SET_KEY_ROW_HEIGHT, 30))
        self.model.set_precision(s.get_int(SET_KEY_CURRENCY_PRECISION, 2))
        self.apply_key_display_settings()
        self.apply_theme(s.get_str(SET_KEY_THEME, "light"))

        # 列宽 / 顺序 / 隐藏
        widths = s.get_dict(SET_KEY_COLUMN_WIDTHS, {})
        header = self.table.horizontalHeader()
        hidden = set(s.get_list(SET_KEY_HIDDEN_COLUMNS, DEFAULT_SETTINGS[SET_KEY_HIDDEN_COLUMNS]))
        for key in DSKM_COLUMN_KEYS:
            logical = DSKM_COLUMN_INDEX.get(key)
            if logical is None:
                continue
            width = widths.get(key)
            if isinstance(width, (int, float)) and int(width) > 20:
                self.table.setColumnWidth(logical, int(width))
            else:
                self.table.setColumnWidth(logical, DSKM_COLUMN_SPECS[logical][3])
            self.table.setColumnHidden(logical, key in hidden)
        order = [k for k in s.get_list(SET_KEY_COLUMN_ORDER, []) if k in DSKM_COLUMN_INDEX]
        for visual_pos, key in enumerate(order):
            logical = DSKM_COLUMN_INDEX[key]
            current = header.visualIndex(logical)
            if current != visual_pos and current >= 0:
                header.moveSection(current, visual_pos)
        self._sanitize_hidden_columns()

        # 排序
        sort_key = s.get_str(SET_KEY_SORT_COLUMN, "")
        sort_order = s.get_str(SET_KEY_SORT_ORDER, "")
        if sort_key in DSKM_COLUMN_INDEX and sort_order in ("asc", "desc"):
            self.model.apply_sort(sort_key, sort_order)
        self._update_sort_indicator()

        # 面板可见性
        stats_visible = s.get_bool(SET_KEY_STATISTICS_BAR, True)
        self.stats_widget.setVisible(stats_visible)
        if hasattr(self, "act_stats"):
            self.act_stats.setChecked(stats_visible)
        log_visible = s.get_bool(SET_KEY_LOG_VISIBLE, True)
        self.log_panel.setVisible(log_visible)
        if hasattr(self, "act_log"):
            self.act_log.setChecked(log_visible)
        if hasattr(self, "act_theme"):
            self.act_theme.setChecked(s.get_str(SET_KEY_THEME, "light") == "dark")
        if hasattr(self, "act_top"):
            self.act_top.setChecked(s.get_bool(SET_KEY_WINDOW_ON_TOP, False))

        sizes = s.get_list(SET_KEY_SPLITTER_SIZES, [])
        if len(sizes) == 2 and all(isinstance(x, int) for x in sizes):
            self.splitter.setSizes(sizes)
        self.log_panel.chk_autoscroll.setChecked(s.get_bool(SET_KEY_LOG_AUTOSCROLL, True))
        self.log_panel.view.setMaximumBlockCount(s.get_int(SET_KEY_LOG_MAX_LINES, 2000))

        # 过滤器
        self.edit_search.blockSignals(True)
        self.edit_search.setText(s.get_str(SET_KEY_SEARCH_TEXT, ""))
        self.edit_search.blockSignals(False)
        self.proxy.set_search_text(self.edit_search.text())
        status_value = s.get_str(SET_KEY_STATUS_FILTER, "all")
        index = self.cmb_status.findData(status_value)
        self.cmb_status.blockSignals(True)
        self.cmb_status.setCurrentIndex(index if index >= 0 else 0)
        self.cmb_status.blockSignals(False)
        self.proxy.set_status_filter(self.cmb_status.currentData())

        self._restart_auto_timer()

    def apply_theme(self, theme: str) -> None:
        app = QApplication.instance()
        if app is not None:
            dskm_apply_dark_palette(app, theme)
            app.setStyleSheet(dskm_build_stylesheet(theme))
        if hasattr(self, "act_theme"):
            self.act_theme.setChecked(theme == "dark")
            self.act_theme.setText("浅色主题" if theme == "dark" else "深色主题")

    def apply_font_size(self, size: int) -> None:
        size = max(7, min(20, int(size or 9)))
        app = QApplication.instance()
        base = dskm_best_ui_font(size)
        if app is not None:
            app.setFont(base)
        self.table.setFont(base)
        self.log_panel.view.setFont(dskm_mono_font(max(7, size - 1)))

    def apply_always_on_top(self, enabled: bool) -> None:
        flags = self.windowFlags()
        if enabled:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        if self.isVisible():
            self.show()
        if hasattr(self, "act_top"):
            self.act_top.setChecked(bool(enabled))

    def _restart_auto_timer(self) -> None:
        enabled = self.settings.get_bool(SET_KEY_AUTO_REFRESH, False)
        interval = max(10, self.settings.get_int(SET_KEY_AUTO_REFRESH_INTERVAL, 30))
        if enabled:
            self._auto_timer.start(interval * 1000)
            DSKM_LOG.info(f"自动刷新已开启：每 {interval} 秒查询一次")
        else:
            self._auto_timer.stop()

    def _on_auto_refresh_tick(self) -> None:
        if self._closing or self.controller.is_running():
            return
        if not self.model.records():
            return
        DSKM_LOG.info("自动刷新触发批量查询")
        self.start_query(force=True, silent=True)

    # ==================================================================
    #  密钥载入 / 增删改
    # ==================================================================
    def reload_keys_from_env(self, initial: bool = False) -> None:
        """从 .env 读取全部密钥，与元数据合并后刷新表格。"""
        if self.controller.is_running():
            self.flash_status("正在查询中，请先停止再重新载入密钥")
            return
        try:
            keys = self.env_store.get_api_keys()
        except Exception as exc:
            DSKM_LOG.error(f"读取 .env 失败：{exc}")
            keys = []

        if self.settings.get_bool(SET_KEY_AUTO_DEDUPE, True):
            deduped: List[str] = []
            seen = set()
            for key in keys:
                if key in seen:
                    continue
                seen.add(key)
                deduped.append(key)
            if len(deduped) != len(keys):
                DSKM_LOG.warn(f"检测到 {len(keys) - len(deduped)} 个重复密钥，已自动去重")
                try:
                    self.env_store.set_api_keys(deduped)
                except Exception:
                    pass
            keys = deduped

        records: List[DSKMKeyRecord] = []
        valid_uids: List[str] = []
        for key in keys:
            uid = dskm_key_uid(key)
            record = DSKMKeyRecord(api_key=key, uid=uid)
            record.apply_meta_dict(self.meta_store.get(uid))
            records.append(record)
            valid_uids.append(uid)

        self.model.set_records(records)
        self._refresh_currency_filter()
        self.update_statistics()
        self.update_window_title()
        if not initial:
            DSKM_LOG.success(f"已从 .env 重新载入 {len(records)} 个密钥")
            self.flash_status(f"已重新载入 {len(records)} 个密钥")
        else:
            DSKM_LOG.info(f"已载入 {len(records)} 个密钥")
        if records and self.settings.get_bool(SET_KEY_CHECK_ON_STARTUP, False):
            QTimer.singleShot(1000, lambda: self.start_query(force=True, silent=True))

    def _persist_keys_to_env(self) -> None:
        keys = [record.api_key for record in self.model.records()]
        try:
            self.env_store.set_api_keys(keys)
        except Exception as exc:
            DSKM_LOG.error(f"写入 .env 失败：{exc}")

    def add_key_dialog(self) -> None:
        existing = [record.api_key for record in self.model.records()]
        dialog = DSKMKeyEditDialog(self._make_client, None, existing, self)
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        key = values["api_key"]
        if any(record.api_key == key for record in self.model.records()):
            self.flash_status("该密钥已存在")
            return
        record = DSKMKeyRecord(api_key=key, uid=dskm_key_uid(key))
        record.alias = values["alias"]
        record.note = values["note"]
        record.enabled = values["enabled"]
        self.model.append_record(record)
        self._persist_keys_to_env()
        self.meta_store.set(record.uid, record.to_meta_dict())
        self._refresh_currency_filter()
        self.update_statistics()
        self.update_window_title()
        self.select_record_uid(record.uid)
        DSKM_LOG.success(f"已添加密钥：{record.display_alias}（{self.model.format_key_text(key)}）")

        test_result = values.get("test_result")
        if isinstance(test_result, DSKMQueryResult):
            record.apply_result(test_result, self.settings.get_int(SET_KEY_HISTORY_LIMIT, 200))
            self.model.refresh_uid(record.uid)
            self.meta_store.set(record.uid, record.to_meta_dict())
            self.update_statistics()
        else:
            self._query_single(record)
        self.flash_status(f"已添加 {record.display_alias}，正在查询余额…")

    def edit_selected_key(self) -> None:
        records = self.selected_records()
        if not records:
            self.flash_status("请先选择要编辑的密钥")
            return
        self.edit_key_record(records[0])

    def edit_key_record(self, record: DSKMKeyRecord) -> None:
        existing = [r.api_key for r in self.model.records() if r.uid != record.uid]
        dialog = DSKMKeyEditDialog(self._make_client, record, existing, self)
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        old_key = record.api_key
        new_key = values["api_key"]
        if new_key != old_key:
            # Key 变了＝换了一条记录：清理旧元数据
            self.meta_store.remove(record.uid)
            record.api_key = new_key
            record.uid = dskm_key_uid(new_key)
            record.last_status = STATUS_PENDING
            record.progress = 0
            record.balance_infos = []
            record.is_available = None
            record.error_message = ""
            record.http_status = 0
            record.history = []
        record.alias = values["alias"]
        record.note = values["note"]
        record.enabled = values["enabled"]
        self.model.refresh_all()
        self._persist_keys_to_env()
        self.meta_store.set(record.uid, record.to_meta_dict())
        self.update_statistics()
        self.update_window_title()
        DSKM_LOG.success(f"已保存密钥：{record.display_alias}")
        if new_key != old_key:
            self._query_single(record)

    def delete_selected_keys(self) -> None:
        records = self.selected_records()
        if not records:
            self.flash_status("请先选择要删除的密钥")
            return
        if self.settings.get_bool(SET_KEY_CONFIRM_ON_DELETE, True):
            names = "、".join(r.display_alias for r in records[:5])
            more = f" 等 {len(records)} 个" if len(records) > 5 else ""
            answer = QMessageBox.question(
                self, "确认删除",
                f"确定要从 .env 中删除 {names}{more} 密钥吗？\n该操作不可撤销。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
        for record in records:
            self.meta_store.remove(record.uid, save=False)
            self.model.remove_uid(record.uid)
        self.meta_store.save()
        self._persist_keys_to_env()
        self._refresh_currency_filter()
        self.update_statistics()
        self.update_window_title()
        DSKM_LOG.warn(f"已删除 {len(records)} 个密钥")
        self.flash_status(f"已删除 {len(records)} 个密钥")

    def clear_all_keys(self) -> None:
        if not self.model.records():
            self.flash_status("当前没有密钥")
            return
        if self.settings.get_bool(SET_KEY_CONFIRM_ON_CLEAR, True):
            answer = QMessageBox.question(
                self, "确认清空",
                f"确定要清空全部 {len(self.model.records())} 个密钥吗？\n"
                "（只清除 .env 中的密钥列表，元数据保留在 keys_meta.json 中）",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
        self.stop_query()
        self.model.clear_all()
        self._persist_keys_to_env()
        self._refresh_currency_filter()
        self.update_statistics()
        self.update_window_title()
        DSKM_LOG.warn("已清空全部密钥")
        self.flash_status("已清空全部密钥")

    def set_selected_enabled(self, enabled: bool) -> None:
        records = self.selected_records()
        if not records:
            self.flash_status("请先选择密钥")
            return
        for record in records:
            record.enabled = bool(enabled)
            if not enabled:
                record.last_status = STATUS_DISABLED
                record.progress = 0
            elif record.last_status == STATUS_DISABLED:
                record.last_status = STATUS_PENDING
                record.progress = 0
            self.meta_store.set(record.uid, record.to_meta_dict(), save=False)
        self.meta_store.save()
        self.model.refresh_all()
        self.update_statistics()
        self.update_action_states()
        DSKM_LOG.info(f"已{'启用' if enabled else '停用'} {len(records)} 个密钥")

    def move_selected(self, delta: int) -> None:
        records = self.selected_records()
        if len(records) != 1:
            self.flash_status("上移/下移一次只能操作一行")
            return
        target = records[0]
        all_records = self.model.records()
        index = next((i for i, r in enumerate(all_records) if r.uid == target.uid), -1)
        if index < 0:
            return
        new_index = index + delta
        if not (0 <= new_index < len(all_records)):
            self.flash_status("已经到边界了")
            return
        all_records[index], all_records[new_index] = all_records[new_index], all_records[index]
        self.model.set_records(all_records)
        self._persist_keys_to_env()
        self.select_record_uid(target.uid)
        self.flash_status("顺序已调整并写入 .env")

    def cleanup_duplicate_keys(self) -> None:
        records = self.model.records()
        seen = set()
        duplicates = []
        for record in records:
            if record.api_key in seen:
                duplicates.append(record)
            else:
                seen.add(record.api_key)
        if not duplicates:
            self.flash_status("没有发现重复密钥")
            DSKM_LOG.info("重复检查完成：没有重复")
            return
        for record in duplicates:
            self.model.remove_uid(record.uid)
        self._persist_keys_to_env()
        self.update_statistics()
        self.update_window_title()
        DSKM_LOG.warn(f"已清理 {len(duplicates)} 个重复密钥")
        self.flash_status(f"已清理 {len(duplicates)} 个重复密钥")

    def import_keys_dialog(self) -> None:
        dialog = DSKMImportDialog(self)
        default_dir = self.settings.get_str(SET_KEY_IMPORT_DIR, "")
        if default_dir:
            pass
        if dialog.exec() != QDialog.Accepted:
            return
        keys = dialog.imported_keys()
        if not keys:
            return
        existed = {record.api_key for record in self.model.records()}
        added = 0
        skipped = 0
        self._begin_busy(f"正在导入 {len(keys)} 个密钥…")
        try:
            for index, key in enumerate(keys):
                if key in existed:
                    skipped += 1
                    continue
                record = DSKMKeyRecord(api_key=key, uid=dskm_key_uid(key))
                record.apply_meta_dict(self.meta_store.get(record.uid))
                self.model.append_record(record)
                existed.add(key)
                added += 1
                if index % 20 == 0:
                    self._busy_progress(int(100 * index / max(1, len(keys))))
            self._persist_keys_to_env()
            self._refresh_currency_filter()
            self.update_statistics()
            self.update_window_title()
        finally:
            self._end_busy(f"导入完成：新增 {added} 个，跳过重复 {skipped} 个")
        DSKM_LOG.success(f"批量导入完成：新增 {added} 个，跳过重复 {skipped} 个")

    # ==================================================================
    #  查询
    # ==================================================================
    def _make_client(self) -> DSKMDeepSeekClient:
        return DSKMDeepSeekClient(
            api_base=self.env_store.get_api_base(),
            timeout=self.settings.get_int(SET_KEY_TIMEOUT, self.env_store.get_int(DSKMEnvStore.KEY_TIMEOUT, 20)),
            max_retry=self.settings.get_int(SET_KEY_MAX_RETRY, self.env_store.get_int(DSKMEnvStore.KEY_MAX_RETRY, 2)),
            verify_ssl=self.settings.get_bool(SET_KEY_VERIFY_SSL, True),
            proxy_url=self.settings.get_str(SET_KEY_PROXY_URL, "") if self.settings.get_bool(SET_KEY_PROXY_ENABLED, False) else "",
        )

    def _rebuild_client(self) -> None:
        self._client = self._make_client()
        self.controller.set_concurrency(
            self.settings.get_int(SET_KEY_CONCURRENCY, self.env_store.get_int(DSKMEnvStore.KEY_CONCURRENCY, 5))
        )

    def _query_targets_for(self, only_selection: bool, force: bool, silent: bool) -> List[DSKMKeyRecord]:
        """确定本次查询范围。

        规则：
          • 停用（enabled=False）的密钥 **永远不参与批量查询**（用户显式停用就是不想查它）；
          • “仅查询所选行”时不再套用“仅查已启用”过滤，方便手工补查单个密钥；
          • force=True（强制重查）只绕过“跳过最近查询”缓存，不改变启用范围。
        """
        if only_selection:
            records = self.selected_records()
        else:
            records = self.model.records()
            if self.settings.get_bool(SET_KEY_ONLY_ENABLED_CHECK, True):
                records = [r for r in records if r.enabled]
        if not force and self.settings.get_bool(SET_KEY_SKIP_RECENT_CHECK, False):
            threshold = max(30, self.settings.get_int(SET_KEY_SKIP_RECENT_SECONDS, 300))
            now = _dt.datetime.now()
            filtered: List[DSKMKeyRecord] = []
            for record in records:
                parsed = dskm_parse_time_text(record.last_checked)
                if parsed is not None and (now - parsed).total_seconds() < threshold:
                    continue
                filtered.append(record)
            skipped = len(records) - len(filtered)
            if skipped and not silent:
                DSKM_LOG.info(f"按设置跳过 {skipped} 个最近查询过的密钥")
            records = filtered
        return records

    def start_query(self, force: bool = False, only_selection: bool = False, silent: bool = False) -> None:
        if self.controller.is_running():
            self.flash_status("已有查询正在进行，请等待完成或点击停止")
            return
        if not DSKM_HAS_REQUESTS:
            QMessageBox.critical(self, "缺少依赖", "未检测到 requests 库，请先执行：\n\npip install requests")
            return
        targets = self._query_targets_for(only_selection, force, silent)
        if not targets:
            self.flash_status("没有需要查询的密钥（可检查“仅查已启用密钥”设置）")
            if not silent:
                DSKM_LOG.warn("没有需要查询的密钥")
            return
        self._rebuild_client()
        self._query_started_at = time.time()
        self._query_targets = len(targets)
        for record in targets:
            record.last_status = STATUS_RUNNING
            record.progress = 3
        self.model.refresh_all()
        self.progress_panel.begin(len(targets), f"正在查询 {len(targets)} 个密钥…")
        jobs = [(record.uid, record.api_key) for record in targets]
        self.controller.start_batch(jobs, self._client)
        self.update_action_states()
        if not silent:
            DSKM_LOG.info(f"开始批量查询：{len(targets)} 个密钥，并发 {self.controller.concurrency()}")

    def _query_single(self, record: DSKMKeyRecord) -> None:
        if not DSKM_HAS_REQUESTS:
            return
        if self.controller.is_running():
            return
        self._rebuild_client()
        record.last_status = STATUS_RUNNING
        record.progress = 3
        self.model.refresh_uid(record.uid)
        self.progress_panel.begin(1, f"正在查询 {record.display_alias}…")
        self.controller.start_batch([(record.uid, record.api_key)], self._client)

    def stop_query(self) -> None:
        if not self.controller.is_running():
            self.flash_status("当前没有正在进行的查询")
            return
        self.controller.cancel()
        self.progress_panel.finish("已请求停止（正在等待已发出的请求结束）")
        DSKM_LOG.warn("用户请求停止批量查询")
        self.update_action_states()
        QTimer.singleShot(600, self._mark_cancelled_records)

    def _mark_cancelled_records(self) -> None:
        changed = False
        for record in self.model.records():
            if record.last_status == STATUS_RUNNING:
                record.last_status = STATUS_PENDING if record.enabled else STATUS_DISABLED
                record.progress = 0
                changed = True
        if changed:
            self.model.refresh_all()
            self.update_statistics()

    def _on_batch_started(self, total: int) -> None:
        self.update_action_states()

    def _on_batch_progress(self, done: int, total: int) -> None:
        self.progress_panel.update_progress(
            done, total, f"正在查询…已完成 {done}/{total}（{int(100 * done / max(1, total))}%）"
        )

    def _on_job_started(self, uid: str) -> None:
        self.model.update_progress(uid, STATUS_RUNNING, 5)

    def _on_job_progress(self, uid: str, pct: int) -> None:
        record = self.model.record_by_uid(uid)
        if record is None or record.last_status != STATUS_RUNNING:
            return
        record.progress = max(5, min(99, int(pct)))
        self.model.refresh_uid(uid)

    def _on_job_finished(self, uid: str, result: object) -> None:
        if not isinstance(result, DSKMQueryResult):
            return
        record = self.model.record_by_uid(uid)
        if record is None:
            return
        history_limit = self.settings.get_int(SET_KEY_HISTORY_LIMIT, 200) if self.settings.get_bool(SET_KEY_KEEP_HISTORY, True) else 0
        record.apply_result(result, history_limit)
        self.model.refresh_uid(uid)
        self.meta_store.set(uid, record.to_meta_dict(), save=False)
        self._pending_meta_save = True
        self._meta_timer.start()
        self._stats_timer.start()
        name = record.display_alias
        if result.ok:
            flag = "可用" if result.is_available else "账户不可用"
            DSKM_LOG.success(f"{name} 查询成功（{flag}，{result.elapsed_ms} ms）：{record.balance_detail_text or '无余额明细'}")
        elif result.error_kind == "cancelled":
            DSKM_LOG.warn(f"{name} 查询已取消")
        else:
            DSKM_LOG.error(f"{name} 查询失败：{result.error_message}")

    def _stats_dirty_pending(self) -> bool:
        return self.controller.pending_count() > 0

    def _on_batch_finished(self, ok: int, fail: int, elapsed: float) -> None:
        self._last_query_result = (ok, fail, elapsed)
        self._stats_timer.stop()
        self._flush_meta_store()
        self._refresh_currency_filter()
        # 排序键可能是状态/余额，查询结束后重新排一次更符合预期
        sort_key, sort_order = self.model.sort_state()
        if sort_order:
            self.model.apply_sort(sort_key, sort_order)
        self.update_statistics()
        self.update_action_states()
        summary = (
            f"查询完成：成功 {ok} 个 / 失败 {fail} 个，"
            f"耗时 {elapsed:.1f} 秒（{dskm_human_seconds(elapsed)}）"
        )
        self.progress_panel.finish(summary)
        self.flash_status(summary)
        DSKM_LOG.info(summary)
        if self.settings.get_bool(SET_KEY_NOTIFY_ON_FINISH, True) and self._query_targets > 1:
            QMessageBox.information(self, "查询完成", summary)

    def _flush_meta_store(self) -> None:
        if not self._pending_meta_save:
            return
        if self.meta_store.save():
            self._pending_meta_save = False
            DSKM_LOG.debug("密钥元数据已保存")
        else:
            DSKM_LOG.warn("密钥元数据保存失败")

    # ==================================================================
    #  统计 / 标题 / 过滤器
    # ==================================================================
    def update_statistics(self) -> None:
        records = self.model.records()
        total = len(records)
        enabled = sum(1 for r in records if r.enabled)
        available = sum(1 for r in records if r.enabled and r.last_status == STATUS_OK and r.is_available)
        unavailable = sum(1 for r in records if r.enabled and (r.last_status == STATUS_FAIL or (r.last_status == STATUS_OK and not r.is_available)))
        pending = sum(1 for r in records if r.enabled and r.last_status in (STATUS_PENDING, STATUS_RUNNING))
        self.card_total.set_value(str(total), f"其中启用 {enabled} 个")
        self.card_enabled.set_value(str(enabled), f"停用 {total - enabled} 个")
        self.card_available.set_value(str(available), "账户状态正常")
        self.card_unavailable.set_value(str(unavailable), "失败或账户停用")
        self.card_pending.set_value(str(pending), f"已完成 {total - pending} 个")
        self.card_unavailable.set_value_color("#B71C1C" if unavailable else "")
        self.card_available.set_value_color("#1B5E20" if available else "")

        # 汇总各币种余额
        sums: Dict[str, Decimal] = {}
        for record in records:
            if not record.enabled or record.last_status != STATUS_OK:
                continue
            for info in record.balance_infos:
                value = dskm_to_decimal(info.total_balance)
                if value is None or not info.currency:
                    continue
                sums[info.currency.upper()] = sums.get(info.currency.upper(), Decimal(0)) + value
        if sums:
            precision = self.settings.get_int(SET_KEY_CURRENCY_PRECISION, 2)
            parts = [f"{dskm_format_amount(v, precision)} {c}" for c, v in sorted(sums.items())]
            self.card_balance.set_value(" + ".join(parts), f"{len(sums)} 个币种")
            self.card_balance.set_value_color("#1B5E20")
        else:
            self.card_balance.set_value("—", "暂无有效余额数据")
            self.card_balance.set_value_color("")

        self.lbl_status_summary.setText(
            f"密钥 {total} 个 ｜ 可用 {available} ｜ 失败 {unavailable} ｜ 待查 {pending}"
        )
        self._refresh_filter_stats()

    def update_window_title(self) -> None:
        total = len(self.model.records())
        sort_key, sort_order = self.model.sort_state()
        suffix = ""
        if sort_order and sort_key in DSKM_COLUMN_TITLES:
            suffix = f" ｜ 排序：{DSKM_COLUMN_TITLES[sort_key]}{'↑' if sort_order == 'asc' else '↓'}"
        self.setWindowTitle(f"{DSKM_APP_NAME}  v{__version__} ｜ {total} 个密钥{suffix}")

    def _refresh_filter_stats(self) -> None:
        shown = self.proxy.rowCount()
        total = self.model.rowCount()
        if shown == total:
            self.lbl_filter_stats.setText(f"显示全部 {total} 行")
        else:
            self.lbl_filter_stats.setText(f"筛选后 {shown} / {total} 行")

    def _refresh_currency_filter(self) -> None:
        currencies: List[str] = []
        for record in self.model.records():
            for info in record.balance_infos:
                if info.currency and info.currency.upper() not in currencies:
                    currencies.append(info.currency.upper())
        current = self.cmb_currency.currentData() or "all"
        self.cmb_currency.blockSignals(True)
        self.cmb_currency.clear()
        self.cmb_currency.addItem("全部", "all")
        for currency in sorted(currencies):
            self.cmb_currency.addItem(currency, currency)
        index = self.cmb_currency.findData(current)
        self.cmb_currency.setCurrentIndex(index if index >= 0 else 0)
        self.cmb_currency.blockSignals(False)
        self.proxy.set_currency_filter(self.cmb_currency.currentData())

    def _on_search_changed(self, text: str) -> None:
        self.proxy.set_search_text(text)
        self.settings.set(SET_KEY_SEARCH_TEXT, text)
        self.schedule_settings_save()
        self._refresh_filter_stats()

    def _on_status_filter_changed(self) -> None:
        value = self.cmb_status.currentData()
        self.proxy.set_status_filter(value)
        self.settings.set(SET_KEY_STATUS_FILTER, value)
        self.schedule_settings_save()
        self._refresh_filter_stats()

    def _on_currency_filter_changed(self) -> None:
        value = self.cmb_currency.currentData()
        self.proxy.set_currency_filter(value)
        self.settings.set(SET_KEY_CURRENCY_FILTER, value)
        self.schedule_settings_save()
        self._refresh_filter_stats()

    def refresh_table_view(self) -> None:
        self.model.refresh_all()
        self._refresh_filter_stats()
        self.update_statistics()
        self.flash_status("表格已刷新")

    def _sanitize_hidden_columns(self) -> None:
        """清掉已不存在的列键（例如程序升级后列定义变化）。"""
        hidden = [k for k in self.settings.get_list(SET_KEY_HIDDEN_COLUMNS, []) if k in DSKM_COLUMN_INDEX]
        visible_count = len(DSKM_COLUMN_KEYS) - len(hidden)
        if visible_count <= 0:
            hidden = list(DEFAULT_SETTINGS[SET_KEY_HIDDEN_COLUMNS])
        self.settings.set(SET_KEY_HIDDEN_COLUMNS, hidden)
        if hidden != self.settings.get_list(SET_KEY_HIDDEN_COLUMNS, []):
            self.schedule_settings_save()

    # ==================================================================
    #  选中与剪贴板
    # ==================================================================
    def selected_records(self) -> List[DSKMKeyRecord]:
        result: List[DSKMKeyRecord] = []
        selection = self.table.selectionModel()
        if selection is None:
            return result
        for index in selection.selectedRows():
            source = self.proxy.mapToSource(index)
            record = self.model.record_at(source.row())
            if record is not None:
                result.append(record)
        return result

    def _copy_text(self, text: str, what: str = "内容") -> None:
        try:
            QApplication.clipboard().setText(text)
            self.flash_status(f"已复制{what}到剪贴板")
            DSKM_LOG.info(f"已复制{what}到剪贴板")
        except Exception as exc:
            DSKM_LOG.warn(f"复制失败：{exc}")

    def select_record_uid(self, uid: str) -> None:
        """按 uid 在（代理）视图中选中并滚动到该行。"""
        source_row = self.model.row_of_uid(uid)
        if source_row < 0:
            return
        source_index = self.model.index(source_row, 0)
        proxy_index = self.proxy.mapFromSource(source_index)
        if not proxy_index.isValid():
            return
        self.table.selectRow(proxy_index.row())
        self.table.scrollTo(proxy_index, QAbstractItemView.PositionAtCenter)

    def copy_selected_rows(self) -> None:
        records = self.selected_records()
        if not records:
            self.flash_status("请先选择要复制的行")
            return
        lines = []
        for record in records:
            lines.append("\t".join([
                record.display_alias,
                record.api_key if self.settings.get_bool(SET_KEY_EXPORT_INCLUDE_KEY, False)
                else self.model.format_key_text(record.api_key),
                "可用" if record.is_available else ("不可用" if record.is_available is False else ""),
                record.total_balance_text,
                record.granted_text,
                record.topped_up_text,
                record.currency_text,
                record.balance_detail_text,
                record.status_text,
                record.last_checked,
                record.error_message,
            ]))
        self._copy_text("\n".join(lines), f"{len(records)} 行信息")

    def copy_table_as_tsv(self) -> None:
        rows: List[str] = []
        visible_keys = [
            key for key in DSKM_COLUMN_KEYS
            if not self.table.isColumnHidden(DSKM_COLUMN_INDEX[key])
        ]
        rows.append("\t".join(DSKM_COLUMN_TITLES.get(k, k) for k in visible_keys))
        for record in self.model.view_records():
            rows.append("\t".join(
                self.model.display_value_of(record, key).replace("\t", " ") for key in visible_keys
            ))
        self._copy_text("\n".join(rows), "表格内容（TSV）")

    # ==================================================================
    #  详情 / 列设置 / 设置 / 帮助
    # ==================================================================
    def show_selected_detail(self) -> None:
        records = self.selected_records()
        if not records:
            self.flash_status("请先选择一行")
            return
        self.show_detail_for(records[0])

    def show_detail_for(self, record: DSKMKeyRecord) -> None:
        dialog = DSKMKeyDetailDialog(
            record,
            self.settings.get_int(SET_KEY_CURRENCY_PRECISION, 2),
            self,
            key_display=self.key_display_config(),
        )
        dialog.exec()

    def configure_columns(self) -> None:
        dialog = DSKMColumnDialog(
            self.settings.get_list(SET_KEY_HIDDEN_COLUMNS, DEFAULT_SETTINGS[SET_KEY_HIDDEN_COLUMNS]),
            self.settings.get_list(SET_KEY_COLUMN_ORDER, list(DSKM_COLUMN_KEYS)),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        hidden = dialog.hidden_columns()
        order = dialog.column_order()
        self.settings.set(SET_KEY_HIDDEN_COLUMNS, hidden)
        self.settings.set(SET_KEY_COLUMN_ORDER, order)
        header = self.table.horizontalHeader()
        for key in DSKM_COLUMN_KEYS:
            logical = DSKM_COLUMN_INDEX.get(key)
            if logical is None:
                continue
            self.table.setColumnHidden(logical, key in hidden)
        for visual_pos, key in enumerate(order):
            logical = DSKM_COLUMN_INDEX.get(key)
            if logical is None:
                continue
            current = header.visualIndex(logical)
            if current != visual_pos and current >= 0:
                header.moveSection(current, visual_pos)
        self._save_settings_now()
        self.flash_status(f"显示列设置已保存（隐藏 {len(hidden)} 列）")
        DSKM_LOG.info(f"显示列设置已更新：隐藏 {len(hidden)} 列")

    def open_settings_dialog(self) -> None:
        if getattr(self, "_settings_dialog", None) is not None:
            self._settings_dialog.refresh_from_store()
            self._settings_dialog.show()
            self._settings_dialog.raise_()
            self._settings_dialog.activateWindow()
            return
        dialog = DSKMSettingsDialog(self.settings, self.env_store, self)
        dialog.setting_changed.connect(self._on_dialog_setting_changed)
        dialog.network_setting_changed.connect(self._on_dialog_network_changed)
        dialog.column_settings_requested.connect(self.configure_columns)
        dialog.env_editor_requested.connect(self.open_env_editor)
        dialog.reset_requested.connect(self.reset_ui_layout)
        dialog.finished.connect(lambda *_: setattr(self, "_settings_dialog", None))
        self._settings_dialog = dialog
        dialog.show()

    def _on_dialog_setting_changed(self, key: str, value: Any) -> None:
        self.set_setting(key, value)

    def _on_dialog_network_changed(self, key: str, value: Any) -> None:
        """网络设置：同时写入 .env（权威）与 settings.json（镜像）。"""
        mapping = {
            SET_KEY_API_BASE: DSKMEnvStore.KEY_API_BASE,
            SET_KEY_TIMEOUT: DSKMEnvStore.KEY_TIMEOUT,
            SET_KEY_MAX_RETRY: DSKMEnvStore.KEY_MAX_RETRY,
            SET_KEY_CONCURRENCY: DSKMEnvStore.KEY_CONCURRENCY,
        }
        env_key = mapping.get(key)
        if env_key:
            try:
                self.env_store.set_values({env_key: str(value)})
            except Exception as exc:
                DSKM_LOG.error(f"写入 .env 失败：{exc}")
        self.set_setting(key, value)
        DSKM_LOG.info(f"网络设置已更新：{key} = {value}")

    def open_env_editor(self) -> None:
        dialog = DSKMEnvEditorDialog(self.env_store, self)
        if dialog.exec() == QDialog.Accepted:
            self.reload_keys_from_env()

    def show_log_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("日志文件内容")
        dialog.resize(900, 620)
        layout = QVBoxLayout(dialog)
        browser = QPlainTextEdit()
        browser.setReadOnly(True)
        browser.setFont(dskm_mono_font(9))
        try:
            path = DSKM_LOG._current_log_file()  # noqa: SLF001 - 内部使用
            text = path.read_text(encoding="utf-8", errors="replace") if path.exists() else "（日志文件尚未生成）"
            browser.setPlainText(text[-200000:])
        except Exception as exc:
            browser.setPlainText(f"读取日志失败：{exc}")
        layout.addWidget(browser, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("关闭")
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def show_help(self) -> None:
        DSKMHelpDialog(self).exec()

    def show_about(self) -> None:
        DSKMAboutDialog(self).exec()

    def reset_ui_layout(self) -> None:
        self.settings.set(SET_KEY_HIDDEN_COLUMNS, list(DEFAULT_SETTINGS[SET_KEY_HIDDEN_COLUMNS]))
        self.settings.set(SET_KEY_COLUMN_WIDTHS, {})
        self.settings.set(SET_KEY_COLUMN_ORDER, list(DSKM_COLUMN_KEYS))
        self.settings.set(SET_KEY_SPLITTER_SIZES, [])
        header = self.table.horizontalHeader()
        for key in DSKM_COLUMN_KEYS:
            logical = DSKM_COLUMN_INDEX[key]
            self.table.setColumnHidden(logical, key in DEFAULT_SETTINGS[SET_KEY_HIDDEN_COLUMNS])
            self.table.setColumnWidth(logical, DSKM_COLUMN_SPECS[logical][3])
            current = header.visualIndex(logical)
            if current != logical and current >= 0:
                header.moveSection(current, logical)
        self.splitter.setSizes([max(200, self.height() - 200), 150])
        self._save_settings_now()
        self.flash_status("界面布局已重置")

    def reset_window_position(self) -> None:
        rects = dskm_screen_available_rects()
        primary = rects[0]
        w = min(1420, max(880, primary.width() - 120))
        h = min(860, max(560, primary.height() - 120))
        x = primary.x() + max(0, (primary.width() - w) // 2)
        y = primary.y() + max(0, (primary.height() - h) // 4)
        if self.isMaximized():
            self.showNormal()
        self.resize(w, h)
        self.move(x, y)
        self._save_settings_now()
        self.flash_status("窗口已居中显示")

    # ==================================================================
    #  视图开关
    # ==================================================================
    def toggle_statistics_bar(self) -> None:
        visible = not self.stats_widget.isVisible()
        self.stats_widget.setVisible(visible)
        self.set_setting(SET_KEY_STATISTICS_BAR, visible)

    def toggle_log_panel(self) -> None:
        visible = not self.log_panel.isVisible()
        self.log_panel.setVisible(visible)
        self.set_setting(SET_KEY_LOG_VISIBLE, visible)

    def toggle_theme(self) -> None:
        theme = "dark" if self.settings.get_str(SET_KEY_THEME, "light") != "dark" else "light"
        self.set_setting(SET_KEY_THEME, theme)
        DSKM_LOG.info(f"主题已切换为{'深色' if theme == 'dark' else '浅色'}")

    def toggle_always_on_top(self) -> None:
        enabled = not self.settings.get_bool(SET_KEY_WINDOW_ON_TOP, False)
        self.set_setting(SET_KEY_WINDOW_ON_TOP, enabled)

    # ==================================================================
    #  导出
    # ==================================================================
    def export_results_dialog(self) -> None:
        if not self.model.records():
            self.flash_status("没有可导出的数据")
            return
        menu = QMenu(self)
        menu.addAction("导出为 CSV（*.csv）", lambda: self.export_results("csv"))
        menu.addAction("导出为 Excel（*.xlsx）", lambda: self.export_results("xlsx"))
        menu.addAction("导出为 JSON（*.json）", lambda: self.export_results("json"))
        menu.addAction("导出为 Markdown 报表（*.md）", lambda: self.export_results("md"))
        menu.exec(self.toolbar.mapToGlobal(QPoint(0, self.toolbar.height())))

    def _export_rows(self) -> List[DSKMKeyRecord]:
        """导出当前显示（筛选 + 排序后）的行。"""
        return self.model.view_records()

    def export_results(self, fmt: str) -> None:
        fmt = (fmt or "csv").lower()
        records = self._export_rows()
        if not records:
            self.flash_status("没有可导出的数据")
            return
        include_key = self.settings.get_bool(SET_KEY_EXPORT_INCLUDE_KEY, False)
        default_dir = self.settings.get_str(SET_KEY_EXPORT_DIR, "") or str(dskm_data_dir())
        ext_map = {"csv": "csv", "xlsx": "xlsx", "json": "json", "md": "md", "markdown": "md", "txt": "txt"}
        ext = ext_map.get(fmt, "csv")
        default_name = str(Path(default_dir) / f"DeepSeek密钥余额_{dskm_now_stamp()}.{ext}")
        filters = {
            "csv": "CSV 文件 (*.csv)",
            "xlsx": "Excel 工作簿 (*.xlsx)",
            "json": "JSON 文件 (*.json)",
            "md": "Markdown 文件 (*.md)",
            "markdown": "Markdown 文件 (*.md)",
            "txt": "文本文件 (*.txt)",
        }
        path, _ = QFileDialog.getSaveFileName(self, "导出查询结果", default_name, filters.get(fmt, "所有文件 (*.*)"))
        if not path:
            return
        self._begin_busy(f"正在导出 {len(records)} 行…")
        try:
            if fmt == "csv":
                self._export_csv(path, records, include_key)
            elif fmt == "xlsx":
                self._export_xlsx(path, records, include_key)
            elif fmt == "json":
                self._export_json(path, records, include_key)
            elif fmt in ("md", "markdown"):
                self._export_markdown(path, records, include_key)
            else:
                self._export_csv(path, records, include_key)
            self.settings.set(SET_KEY_EXPORT_DIR, str(Path(path).parent))
            self.settings.set(SET_KEY_LAST_EXPORT_FORMAT, fmt)
            self._save_settings_now()
            DSKM_LOG.success(f"导出成功：{path}（{len(records)} 行，{'含' if include_key else '不含'}完整 Key）")
            self._end_busy(f"导出完成：{path}")
            self.flash_status(f"导出完成：{path}")
        except Exception as exc:
            self._end_busy(f"导出失败：{exc}")
            DSKM_LOG.error(f"导出失败：{exc}")
            QMessageBox.critical(self, "导出失败", f"{type(exc).__name__}: {exc}")

    def _export_headers(self, include_key: bool) -> List[str]:
        headers = ["名称", "备注"]
        headers.append("API Key" if include_key else "API Key（掩码）")
        headers.extend([
            "账户是否可用 is_available",
            "余额明细 balance_infos",
            "总可用余额 total_balance",
            "赠金余额 granted_balance",
            "充值余额 topped_up_balance",
            "货币单位 currency",
            "状态",
            "最后检查时间",
            "耗时(ms)",
            "HTTP 状态码",
            "累计查询次数",
            "成功次数",
            "失败次数",
            "错误 / 说明",
            "是否启用",
        ])
        return headers

    def _export_row_values(self, record: DSKMKeyRecord, include_key: bool) -> List[str]:
        return [
            record.display_alias,
            record.note,
            record.api_key if include_key else self.model.format_key_text(record.api_key),
            "" if record.is_available is None else ("可用" if record.is_available else "不可用"),
            record.balance_detail_text,
            record.total_balance_text,
            record.granted_text,
            record.topped_up_text,
            record.currency_text,
            record.status_text,
            record.last_checked,
            str(record.elapsed_ms),
            str(record.http_status),
            str(record.check_count),
            str(record.success_count),
            str(record.fail_count),
            record.error_message,
            "是" if record.enabled else "否",
        ]

    def _export_csv(self, path: str, records: Sequence[DSKMKeyRecord], include_key: bool) -> None:
        with open(path, "w", encoding="utf-8-sig", newline="") as fp:
            writer = csv.writer(fp)
            writer.writerow(self._export_headers(include_key))
            for record in records:
                writer.writerow(self._export_row_values(record, include_key))

    def _export_json(self, path: str, records: Sequence[DSKMKeyRecord], include_key: bool) -> None:
        payload = {
            "导出时间": dskm_now_text(),
            "程序版本": __version__,
            "接口地址": self.env_store.get_api_base() + DSKM_BALANCE_PATH,
            "密钥数量": len(records),
            "包含完整Key": bool(include_key),
            "数据": [],
        }
        for record in records:
            item = {
                "名称": record.display_alias,
                "备注": record.note,
                "API Key": record.api_key if include_key else self.model.format_key_text(record.api_key),
                "is_available": record.is_available,
                "balance_infos": [i.to_dict() for i in record.balance_infos],
                "total_balance": record.primary_balance.total_balance if record.primary_balance else "",
                "granted_balance": record.primary_balance.granted_balance if record.primary_balance else "",
                "topped_up_balance": record.primary_balance.topped_up_balance if record.primary_balance else "",
                "currency": record.primary_balance.currency if record.primary_balance else "",
                "状态": record.status_text,
                "最后检查时间": record.last_checked,
                "耗时ms": record.elapsed_ms,
                "http_status": record.http_status,
                "累计查询次数": record.check_count,
                "是否启用": record.enabled,
                "错误说明": record.error_message,
                "余额变化历史": record.history,
            }
            payload["数据"].append(item)
        with open(path, "w", encoding="utf-8") as fp:
            json.dump(payload, fp, ensure_ascii=False, indent=2)

    def _export_xlsx(self, path: str, records: Sequence[DSKMKeyRecord], include_key: bool) -> None:
        if not DSKM_HAS_OPENPYXL:
            DSKM_LOG.warn("未安装 openpyxl，已自动改用 CSV 导出")
            self._export_csv(path.rsplit(".", 1)[0] + ".csv", records, include_key)
            return
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "密钥余额"
        headers = self._export_headers(include_key)
        sheet.append(headers)
        for record in records:
            sheet.append(self._export_row_values(record, include_key))
        for idx, header in enumerate(headers, start=1):
            column = openpyxl.utils.get_column_letter(idx)
            width = max(12, min(42, len(str(header)) + 6))
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="ECEFF1")
        sheet.auto_filter.ref = sheet.dimensions
        workbook.save(path)

    def _export_markdown(self, path: str, records: Sequence[DSKMKeyRecord], include_key: bool) -> None:
        precision = self.settings.get_int(SET_KEY_CURRENCY_PRECISION, 2)
        sums: Dict[str, Decimal] = {}
        for record in records:
            if record.last_status != STATUS_OK:
                continue
            for info in record.balance_infos:
                value = dskm_to_decimal(info.total_balance)
                if value is None or not info.currency:
                    continue
                sums[info.currency.upper()] = sums.get(info.currency.upper(), Decimal(0)) + value
        lines = [
            "# DeepSeek 密钥余额报表",
            "",
            f"- 生成时间：{dskm_now_text()}",
            f"- 程序版本：v{__version__}",
            f"- 接口地址：`{self.env_store.get_api_base()}{DSKM_BALANCE_PATH}`",
            f"- 密钥数量：{len(records)}",
            f"- 包含完整 Key：{'是' if include_key else '否'}",
        ]
        if sums:
            lines.append("- 余额合计：" + " / ".join(
                f"**{dskm_format_amount(v, precision)} {c}**" for c, v in sorted(sums.items())
            ))
        lines.extend([
            "",
            "| # | 名称 | API Key | 账户是否可用 | 总可用余额 | 赠金余额 | 充值余额 | 货币 | 状态 | 最后检查 |",
            "|---|---|---|---|---|---|---|---|---|---|",
        ])
        for idx, record in enumerate(records, start=1):
            key_text = record.api_key if include_key else self.model.format_key_text(record.api_key)
            available = "—" if record.is_available is None else ("✅ 是" if record.is_available else "❌ 否")
            lines.append(
                f"| {idx} | {record.display_alias} | `{key_text}` | {available} | "
                f"{record.total_balance_text or '—'} | {record.granted_text or '—'} | "
                f"{record.topped_up_text or '—'} | {record.currency_text or '—'} | "
                f"{record.status_text} | {record.last_checked or '—'} |"
            )
        lines.append("")
        with open(path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))

    # ==================================================================
    #  杂项
    # ==================================================================
    def flash_status(self, text: str, timeout_ms: int = 5000) -> None:
        try:
            self.lbl_status_main.setText(text)
        except Exception:
            pass

    def _begin_busy(self, text: str) -> None:
        self.progress_panel.progress.setRange(0, 0)
        self.progress_panel.lbl_status.setText(text)
        self.progress_panel.lbl_count.setText("")
        self.progress_panel.btn_cancel.setEnabled(False)
        QApplication.setOverrideCursor(DSKM_CURSOR_WAIT)

    def _busy_progress(self, percent: int) -> None:
        if self.progress_panel.progress.maximum() == 0:
            self.progress_panel.progress.setRange(0, 100)
        self.progress_panel.progress.setValue(max(0, min(100, int(percent))))

    def _end_busy(self, text: str) -> None:
        while True:
            try:
                QApplication.restoreOverrideCursor()
                if QApplication.overrideCursor() is None:
                    break
            except Exception:
                break
        self.progress_panel.progress.setRange(0, 1000)
        self.progress_panel.progress.setValue(1000)
        self.progress_panel.lbl_status.setText(text)
        self.progress_panel.lbl_count.setText("")

    def _open_path(self, path: Path) -> None:
        try:
            os.startfile(str(path))  # noqa: S606
        except Exception as exc:
            DSKM_LOG.warn(f"打开目录失败：{exc}")
            self.flash_status(f"打开目录失败：{exc}")

    def _open_url(self, url: str) -> None:
        try:
            os.startfile(url)  # noqa: S606
        except Exception:
            try:
                import webbrowser
                webbrowser.open(url)
            except Exception as exc:
                DSKM_LOG.warn(f"打开链接失败：{exc}")

    def update_action_states(self) -> None:
        running = self.controller.is_running()
        has_records = bool(self.model.records())
        has_selection = bool(self.selected_records())
        self.progress_panel.btn_cancel.setEnabled(running)
        for action in (self.act_add, self.act_import, self.act_export, self.act_settings):
            action.setEnabled(True)
        self.act_stop.setEnabled(running)
        self.act_query.setEnabled(not running and has_records)
        self.act_query_force.setEnabled(not running and has_records)
        if hasattr(self, "toolbar"):
            self.lbl_toolbar_hint.setText(
                f"  查询进行中…剩余 {self.controller.pending_count()} 个  "
                if running else "  提示：双击行可查看详情｜右键可复制/编辑/删除  "
            )

    # ==================================================================
    #  窗口事件
    # ==================================================================
    def showEvent(self, event) -> None:  # noqa: D102, N802
        super().showEvent(event)
        if getattr(self, "_first_shown", False):
            return
        self._first_shown = True
        if getattr(self, "_start_maximized", False):
            self.showMaximized()
        QTimer.singleShot(0, lambda: dskm_ensure_window_on_screen(self))

    def closeEvent(self, event) -> None:  # noqa: D102, N802
        if self._closing:
            super().closeEvent(event)
            return
        self._closing = True
        try:
            self._auto_timer.stop()
            self._meta_timer.stop()
            self._save_timer.stop()
            if self.controller.is_running():
                self.controller.cancel()
                self.controller.wait_all(5000)
            self._flush_meta_store()
            self._save_settings_now()
        except Exception as exc:
            DSKM_LOG.warn(f"退出清理时出现问题：{exc}")
        DSKM_LOG.info("程序已退出")
        super().closeEvent(event)


# ============================================================================
#  区域 15：程序入口
# ============================================================================


def dskm_setup_high_dpi() -> None:
    """高分屏适配；必须在创建 QApplication 之前调用。

    Qt6 里高 DPI 缩放**默认就是开的**，`AA_EnableHighDpiScaling` /
    `AA_UseHighDpiPixmaps` 已标记废弃（实测会刷 DeprecationWarning），
    所以这里不再去设它们，只设舍入策略（避免非整数缩放比下界面发虚/错位）。
    """
    try:
        from PySide6.QtCore import Qt as _Qt
        policy = getattr(_Qt, "HighDpiScaleFactorRoundingPolicy", None)
        if policy is not None:
            QGuiApplication.setHighDpiScaleFactorRoundingPolicy(policy.PassThrough)
    except Exception:
        pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    dskm_setup_high_dpi()
    app = QApplication(list(argv) if argv is not None else sys.argv)
    app.setApplicationName(DSKM_APP_NAME)
    app.setApplicationDisplayName(DSKM_APP_NAME)
    app.setApplicationVersion(__version__)
    app.setOrganizationName(DSKM_ORG_NAME)
    app.setQuitOnLastWindowClosed(True)
    try:
        app.setStyle("Fusion")
    except Exception:
        pass
    icon = dskm_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)

    if not DSKM_HAS_REQUESTS:
        QMessageBox.critical(
            None, "缺少依赖",
            "未检测到 requests 库，程序无法查询余额。\n\n请执行：\n    pip install requests\n"
        )

    window = DSKMMainWindow()
    if getattr(window, "_start_maximized", False):
        window.showMaximized()
    else:
        window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
