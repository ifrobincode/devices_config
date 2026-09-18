#!/usr/bin/env python3
# 【逐行说明】指定源文件编码为 UTF-8，保证中文注释和字符串正常处理。
# -*- coding: utf-8 -*-
"""
Author: ifrobincode
Created: 2026-09-12
Version: v8.0
Updated: 2026-09-18
Description:
    基于 Excel + Jinja2 模板的网络设备批量配置工具。
    支持 Preview 和 Deploy 两种运行模式。

Usage:
    1. 将本脚本、devices_info.xlsx、config_template.txt 放在同一目录。
    2. 运行：
           python devices_config_v8.py
    3. 根据提示输入 Excel 和模板文件名。
       文件扩展名可省略，程序会自动补全。
    4. 选择 Preview 或 Deploy。

注意：
    - Preview 模式不会连接任何网络设备，只负责 PRE-CHECK、模板渲染和生成配置文件。
    - Deploy 模式会在 Preview/渲染阶段全部检查通过后，再连接设备并下发配置。
    - Sheet1 的 secret 为可选字段；其他必要登录字段不能为空。
    - Sheet2 中除 host 外的配置变量默认全部不能为空。
"""

# 【逐行说明】允许类型注解中的前向引用，并兼容较新的类型提示写法。
from __future__ import annotations

# 【逐行说明】导入 logging，用于记录控制台和文件日志。
import logging
# 【逐行说明】导入 re，用于正则表达式匹配模板残留和 CLI 错误信息。
import re
# 【逐行说明】导入 sys，用于访问标准输出等系统对象。
import sys
# 【逐行说明】导入线程池和 Future 工具，用于并发下发多台设备配置。
from concurrent.futures import ThreadPoolExecutor, as_completed
# 【逐行说明】导入 dataclass，用于定义结构化数据对象；field 用于定义默认工厂。
from dataclasses import dataclass, field
# 【逐行说明】导入 datetime，用于生成任务输出目录的时间戳。
from datetime import datetime
# 【逐行说明】导入 Path，用于跨平台处理文件和目录路径。
from pathlib import Path
# 【逐行说明】导入 Any，用于表示连接参数中可能存在的多种数据类型。
from typing import Any

# 【逐行说明】导入 pandas，用于读取和处理 Excel 数据。
import pandas as pd
# 【逐行说明】导入 Jinja2 模板环境、严格未定义变量模式、模板语法异常和变量分析工具。
from jinja2 import Environment, StrictUndefined, TemplateSyntaxError, meta
# 【逐行说明】导入 Netmiko 的 ConnectHandler，用于建立网络设备连接。
from netmiko import ConnectHandler
# 【逐行说明】开始导入 Netmiko 的异常类型。
from netmiko.exceptions import (
# 【逐行说明】当前代码行执行本步骤的具体操作。
    AuthenticationException,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    ConfigInvalidException,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    ConnectionException,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    NetmikoTimeoutException,
# 【逐行说明】结束 Netmiko 异常类型导入。
)
# 【修改说明】导入 Paramiko 的 SSHException，用于单独识别 SSH 协议握手阶段的异常。
from paramiko.ssh_exception import SSHException

# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# 全局配置
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================

# 【修改说明】根据程序运行方式确定文件基准目录，确保普通 Python 和 PyInstaller EXE 都从正确位置读取外部文件。
# 【修改说明】判断程序是否运行在 PyInstaller 打包后的 EXE 环境中，以便正确确定用户实际放置 EXE 的目录。
if getattr(sys, "frozen", False):
    # 【修改说明】EXE 模式下使用 sys.executable 获取实际启动的 EXE 路径，避免 __file__ 指向 PyInstaller 的 _MEI 临时目录。
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    # 【修改说明】普通 Python 脚本模式下继续使用源代码文件所在目录作为程序工作目录。
    SCRIPT_DIR = Path(__file__).resolve().parent
# 【逐行说明】定义 Excel 默认文件名。
DEFAULT_EXCEL_NAME = "devices_info.xlsx"
# 【逐行说明】定义 Jinja2 配置模板默认文件名。
DEFAULT_TEMPLATE_NAME = "config_template.txt"

# 【逐行说明】定义设备登录信息所在 Sheet 的名称。
SHEET1_NAME = "Sheet1"
# 【逐行说明】定义配置变量所在 Sheet 的名称。
SHEET2_NAME = "Sheet2"

# 【逐行说明】定义 Excel 文件的扩展名。
EXCEL_EXTENSION = ".xlsx"
# 【逐行说明】定义配置模板文件的扩展名。
TEMPLATE_EXTENSION = ".txt"

# 【逐行说明】定义最多同时运行的设备线程数。
MAX_WORKERS = 20

# 【逐行说明】说明下面是 Sheet1 的字段规则配置。
# Sheet1 字段规则：
# 【逐行说明】True 表示该字段为必要字段，不能为空。
# True  = 必要字段，不能为空。
# 【逐行说明】False 表示该字段为可选字段，可以为空。
# False = 可选字段，可以为空。
# 【逐行说明】创建 Sheet1 字段规则字典，并明确其键为字符串、值为布尔值。
SHEET1_FIELD_RULES: dict[str, bool] = {
# 【逐行说明】host 是设备唯一标识，因此不能为空。
    "host": True,
# 【逐行说明】ip 是设备管理地址，因此不能为空。
    "ip": True,
# 【逐行说明】username 是登录用户名，因此不能为空。
    "username": True,
# 【逐行说明】password 是登录密码，因此不能为空。
    "password": True,
# 【逐行说明】secret 是可选的特权/配置模式密码，因此允许为空。
    "secret": False,
# 【逐行说明】device_type 决定 Netmiko 使用哪个设备平台驱动，因此不能为空。
    "device_type": True,
# 【修改说明】port 决定 Netmiko 实际连接设备使用的 TCP 端口，因此不能为空。
    "port": True,
# 【逐行说明】结束 Sheet1 字段规则字典。
}

# 【逐行说明】说明 Sheet2 只有 host 是关联键，其余列都是配置变量。
# Sheet2 中只有 host 是关联键；其余列均视为配置变量。
# 【逐行说明】定义 Sheet2 的关联键名称。
SHEET2_KEY_FIELD = "host"

# 【逐行说明】说明下面的正则表达式用于发现渲染结果中残留的 Jinja2 标记。
# 用于判断最终渲染结果中是否残留 Jinja2 模板标记。
# 【逐行说明】编译匹配 {{...}}、{%...%}、{#...#} 等未渲染模板标记的正则表达式。
UNRENDERED_TEMPLATE_PATTERN = re.compile(r"({[{%#]).*?([}%#]})", re.DOTALL)

# 【逐行说明】说明下面定义常见的设备 CLI 配置错误关键字。
# 常见 CLI 错误关键字。
# 【逐行说明】提醒：不同厂商和软件版本的错误输出不同，后续可继续扩充。
# 注意：不同厂商/版本输出不同，后续可以继续扩充。
# 【逐行说明】创建不可变的错误关键字正则表达式元组。
CONFIG_ERROR_PATTERNS: tuple[str, ...] = (
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"% ?Invalid input",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"% ?Incomplete command",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"% ?Ambiguous command",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"% ?Error",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Error:",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"ERROR:",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Invalid input",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Invalid command",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Unknown command",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Unrecognized command",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Command not found",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Failure",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Failed",
# 【逐行说明】当前代码行执行本步骤的具体操作。
    r"Error",
# 【逐行说明】结束 CLI 错误关键字定义。
)

# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# 数据结构
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【逐行说明】使用 dataclass 自动生成该数据类的初始化和表示等方法。
@dataclass
# 【逐行说明】定义 PRE-CHECK 检查问题的数据结构。
class ValidationIssue:
    """表示一次 PRE-CHECK 检查结果。"""

    level: str
    message: str


@dataclass
class Device:
    """保存单台设备的登录信息、配置变量和最终配置。"""

# 【逐行说明】保存设备主机名/唯一标识 host。
    host: str
# 【逐行说明】保存传给 Netmiko ConnectHandler 的连接参数。
    connection_params: dict[str, Any]
# 【逐行说明】保存该设备对应的 Jinja2 模板变量。
    variables: dict[str, Any]
# 【逐行说明】保存 Jinja2 渲染后的最终配置文本，默认为空字符串。
    rendered_config: str = ""


# 【逐行说明】使用 dataclass 定义单台设备 Deploy 结果的数据结构。
@dataclass
# 【逐行说明】定义单台设备的 Deploy 执行结果。
class DeploymentResult:
    """保存单台设备的 Deploy 执行结果。"""

    host: str
    status: str
    message: str
    config_file: str = ""


@dataclass
class ValidationContext:
    """保存整个 PRE-CHECK 阶段的数据。"""

# 【逐行说明】保存用户选择的 Excel 文件路径。
    excel_path: Path
# 【逐行说明】保存用户选择的配置模板路径。
    template_path: Path
# 【逐行说明】保存清洗后的 Sheet1 DataFrame。
    sheet1: pd.DataFrame
# 【逐行说明】保存清洗后的 Sheet2 DataFrame。
    sheet2: pd.DataFrame
# 【逐行说明】保存配置模板原始文本。
    template_text: str
# 【逐行说明】保存从 Jinja2 模板中解析出的变量名称集合。
    template_variables: set[str] = field(default_factory=set)
# 【逐行说明】保存经过检查并成功渲染的设备对象列表。
    devices: list[Device] = field(default_factory=list)
# 【逐行说明】保存整个 PRE-CHECK 期间收集到的 ERROR/WARNING。
    issues: list[ValidationIssue] = field(default_factory=list)


# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# 日志
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【修改说明】定义支持原样输出标题的日志格式化器，使标题不带日期、时间和日志级别前缀。
class LogFormatter(logging.Formatter):
    """根据日志记录的 raw 标记决定是否省略标准日志前缀。"""

    # 【修改说明】重写日志格式化方法，仅对明确标记为 raw 的标题记录取消统一前缀。
    def format(self, record: logging.LogRecord) -> str:
        # 【修改说明】检查当前日志记录是否要求原样输出。
        if getattr(record, "raw", False):
            # 【修改说明】直接返回日志消息正文，使标题保持用户指定的纯文本格式。
            return record.getMessage()
        # 【修改说明】其他普通日志继续使用日期、时间和日志级别前缀。
        return super().format(record)


# 【逐行说明】定义日志初始化函数，并声明返回 logging.Logger。
def setup_logger(
    log_file: Path | None = None,
    console_output: bool = True,
) -> logging.Logger:
    """创建控制台日志，并按需增加文件日志。"""
# 【修改说明】允许调用方控制是否输出到 CMD，使 PRE-CHECK 日志落盘时不会重复显示已经输出过的问题。
    logger = logging.getLogger("devices_config")
# 【修改说明】保留 INFO 级别，使 INFO、WARNING 和 ERROR 都可以按日志处理流程输出。
    logger.setLevel(logging.INFO)
# 【修改说明】切换日志输出目标前关闭已有 Handler，避免 Windows 下旧日志文件保持打开状态。
    for handler in logger.handlers:
        # 【修改说明】释放已有 Handler 占用的文件或控制台资源。
        handler.close()
    # 【修改说明】清除旧 Handler，避免同一条日志同时写入旧阶段日志文件。
    logger.handlers.clear()

# 【修改说明】统一定义日志时间、级别和消息格式，保持 CMD 与日志文件的基本格式一致。
    # 【修改说明】使用自定义格式化器，使普通日志保留时间和级别，而标题日志可以原样输出。
    formatter = LogFormatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 【修改说明】关闭 Paramiko 的底层错误日志输出，避免已由业务层捕获处理的 SSHException 再打印完整 traceback。
    logging.getLogger("paramiko").setLevel(logging.CRITICAL)
    # 【修改说明】进一步限制 Paramiko Transport 子日志，确保 SSH banner 异常不会绕过主程序日志格式输出到 CMD。
    logging.getLogger("paramiko.transport").setLevel(logging.CRITICAL)

# 【修改说明】仅在调用方要求 CMD 输出时创建控制台 Handler。
    if console_output:
        # 【修改说明】将日志发送到标准输出，使用户可以实时看到设备执行结果。
        console_handler = logging.StreamHandler(sys.stdout)
        # 【修改说明】为控制台 Handler 应用统一日志格式。
        console_handler.setFormatter(formatter)
        # 【修改说明】将控制台 Handler 注册到 Logger。
        logger.addHandler(console_handler)

# 【修改说明】只有传入文件路径时才创建 FileHandler，从而继续支持日志文件的惰性创建。
    if log_file is not None:
        # 【修改说明】使用 UTF-8 打开阶段日志文件，确保中文日志可以正常保存。
        file_handler = logging.FileHandler(
            log_file,
            encoding="utf-8",
        )
        # 【修改说明】为文件 Handler 应用与 CMD 相同的日志格式。
        file_handler.setFormatter(formatter)
        # 【修改说明】将文件 Handler 注册到 Logger，使指定日志写入文件。
        logger.addHandler(file_handler)

# 【修改说明】返回当前阶段已经配置完成的 Logger。
    return logger


# ============================================================================
# 用户输入
# ============================================================================


def normalize_filename(
    user_input: str,
    default_name: str,
    extension: str,
) -> str:
    """
# 【逐行说明】当前代码行执行本步骤的具体操作。
    规范化用户输入的文件名。

# 【逐行说明】当前代码行执行本步骤的具体操作。
    支持：
# 【逐行说明】当前代码行执行本步骤的具体操作。
        devices_info_switch
# 【逐行说明】当前代码行执行本步骤的具体操作。
        devices_info_switch.xlsx

# 【逐行说明】当前代码行执行本步骤的具体操作。
    两种输入方式。

# 【逐行说明】当前代码行执行本步骤的具体操作。
    如果用户直接输入正确扩展名，不会重复添加扩展名。
    """
    value = user_input.strip()

    if not value:
        return default_name

    if value.lower().endswith(extension.lower()):
        return value

    return f"{value}{extension}"


def prompt_for_files() -> tuple[Path, Path]:
    """让用户输入 Excel 和模板文件名，并检查文件是否存在。"""
# 【逐行说明】开始循环，直到用户输入一个存在的 Excel 文件。
    while True:
# 【逐行说明】读取用户输入的 Excel 文件名。
        excel_input = input(
# 【逐行说明】显示 Excel 默认文件名提示。
            f"请输入设备信息文件名（默认：{DEFAULT_EXCEL_NAME}）："
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )
# 【逐行说明】调用文件名规范化函数处理 Excel 输入。
        excel_name = normalize_filename(
# 【逐行说明】当前代码行执行本步骤的具体操作。
            excel_input,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            DEFAULT_EXCEL_NAME,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            EXCEL_EXTENSION,
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )
# 【逐行说明】将脚本目录和用户输入的文件名拼接为完整路径。
        excel_path = SCRIPT_DIR / excel_name

# 【逐行说明】判断 Excel 路径是否指向一个普通文件。
        if excel_path.is_file():
# 【逐行说明】文件存在时结束 Excel 文件输入循环。
            break

# 【逐行说明】告知用户 Excel 文件不存在。
        print(f"[ERROR] Excel 文件不存在：{excel_path}")
# 【逐行说明】提示用户重新输入。
        print("请重新输入。\n")

# 【逐行说明】开始循环，直到用户输入一个存在的模板文件。
    while True:
# 【逐行说明】读取用户输入的模板文件名。
        template_input = input(
# 【逐行说明】显示模板默认文件名提示。
            f"请输入配置模板文件名（默认：{DEFAULT_TEMPLATE_NAME}）："
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )
# 【逐行说明】调用文件名规范化函数处理模板输入。
        template_name = normalize_filename(
# 【逐行说明】当前代码行执行本步骤的具体操作。
            template_input,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            DEFAULT_TEMPLATE_NAME,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            TEMPLATE_EXTENSION,
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )
# 【逐行说明】将脚本目录和模板文件名拼接为完整路径。
        template_path = SCRIPT_DIR / template_name

# 【逐行说明】判断模板路径是否指向一个普通文件。
        if template_path.is_file():
# 【逐行说明】文件存在时结束模板文件输入循环。
            break

# 【逐行说明】告知用户模板文件不存在。
        print(f"[ERROR] 配置模板文件不存在：{template_path}")
# 【逐行说明】提示用户重新输入。
        print("请重新输入。\n")

# 【逐行说明】返回 Excel 和模板文件的完整路径。
    return excel_path, template_path


# 【逐行说明】定义运行模式选择函数。
def prompt_for_mode() -> str:
    """让用户选择 Preview 或 Deploy。"""
    while True:
        print()
        print("请选择运行模式：")
        print("1. Preview - 只生成配置文件，不连接设备")
        print("2. Deploy  - 生成配置文件并下发设备")

        choice = input("请输入（1/2）：").strip()

        if choice == "1":
            return "preview"

        if choice == "2":
            return "deploy"

        print("[ERROR] 无效选择，请输入 1 或 2。\n")


def prompt_for_deploy_confirmation() -> bool:
    """要求用户输入严格匹配的 YES 或 no 决定是否进入设备操作阶段。"""
# 【逐行说明】输出空行。
    print()
# 【逐行说明】输出分隔线。
    print("=" * 72)
# 【逐行说明】告知用户 PRE-CHECK 已通过。
    print("PRE-CHECK 已通过，配置文件已经生成。")
# 【逐行说明】明确提示下一步将连接设备并下发配置。
    print("接下来将开始连接网络设备并下发配置。")
# 【逐行说明】提醒用户先检查已经生成的配置文件。
    print("请确认已经检查生成的配置文件。")
# 【逐行说明】输出结束分隔线。
    print("=" * 72)

# 【修改说明】持续等待用户输入严格匹配的 YES 或 no，避免其他输入被误认为取消操作。
    while True:
# 【修改说明】提示用户使用规定的大小写输入确认或取消 Deploy。
        confirmation = input("请输入 YES 开始下发配置，输入 no 终止程序：").strip()

# 【修改说明】只有完全匹配大写 YES 才允许进入设备连接和配置下发阶段。
        if confirmation == "YES":
# 【修改说明】返回 True，通知主流程开始执行 Deploy。
            return True

# 【修改说明】只有完全匹配小写 no 才允许用户主动取消 Deploy。
        if confirmation == "no":
# 【修改说明】返回 False，通知主流程终止 Deploy 且不连接任何设备。
            return False

# 【修改说明】其他任何输入都不产生动作，继续等待用户明确输入 YES 或 no。
        print("[ERROR] 输入无效，只能输入大写 YES 或小写 no。")


# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# Excel 读取
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【逐行说明】定义 Excel 加载函数。
def load_excel(excel_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """读取 Excel 的 Sheet1 和 Sheet2。"""
    workbook = pd.ExcelFile(excel_path)

    required_sheets = {SHEET1_NAME, SHEET2_NAME}
    missing_sheets = required_sheets.difference(workbook.sheet_names)

    if missing_sheets:
        raise ValueError(
            f"Excel 缺少必要 Sheet：{', '.join(sorted(missing_sheets))}"
        )

    sheet1 = pd.read_excel(workbook, sheet_name=SHEET1_NAME, dtype=object)
    sheet2 = pd.read_excel(workbook, sheet_name=SHEET2_NAME, dtype=object)

    return sheet1, sheet2


def clean_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """统一处理 Excel 单元格中的 NaN 和字符串首尾空格。"""
# 【逐行说明】复制原 DataFrame，避免直接修改原始对象。
    result = dataframe.copy()

# 【逐行说明】遍历 DataFrame 的每一列。
    for column in result.columns:
# 【逐行说明】对当前列的每个单元格执行统一转换。
        result[column] = result[column].map(
# 【逐行说明】NaN 转为空字符串，否则转换为字符串并去掉两端空格。
            lambda value: "" if pd.isna(value) else str(value).strip()
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )

# 【逐行说明】返回清洗后的 DataFrame。
    return result


# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# PRE-CHECK：Excel 结构
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【逐行说明】定义向 PRE-CHECK 上下文追加问题的辅助函数。
def add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    context: ValidationContext,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    level: str,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    message: str,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> None:
    """向 PRE-CHECK 结果中增加一条 INFO/WARNING/ERROR。"""
    context.issues.append(
        ValidationIssue(
            level=level,
            message=message,
        )
    )


def check_required_columns(
    context: ValidationContext,
) -> None:
    """检查 Sheet1 和 Sheet2 是否存在必要字段。"""
    # 【修改说明】将 Sheet1 列名转换成集合，便于判断必要字段是否存在。
    sheet1_columns = set(context.sheet1.columns)
    # 【修改说明】将 Sheet2 列名转换成集合，便于判断 host 是否存在。
    sheet2_columns = set(context.sheet2.columns)

    # 【修改说明】如果 Sheet1 不存在任何列，说明实际 Sheet1 缺失，后续不再重复报告每个字段缺失。
    if context.sheet1.empty and len(context.sheet1.columns) == 0:
        # 【修改说明】记录 Sheet1 缺失这一条根因错误，避免产生大量无意义的字段错误。
        add_issue(
            context,
            "ERROR",
            f"Excel 缺少必要 Sheet：{SHEET1_NAME}",
        )
        # 【修改说明】Sheet1 缺失时直接结束本函数，避免继续检查不存在的字段。
        return

    # 【修改说明】如果 Sheet2 没有任何列，说明实际 Sheet2 缺失，后续不再重复报告 host 缺失。
    if context.sheet2.empty and len(context.sheet2.columns) == 0:
        # 【修改说明】记录 Sheet2 缺失这一条根因错误，避免产生大量无意义的字段错误。
        add_issue(
            context,
            "ERROR",
            f"Excel 缺少必要 Sheet：{SHEET2_NAME}",
        )
        # 【修改说明】Sheet2 缺失时直接结束本函数，避免继续检查不存在的字段。
        return

    # 【修改说明】遍历 Sheet1 的字段规则，并只对 required=True 的字段执行缺失检查。
    for field_name, required in SHEET1_FIELD_RULES.items():
        # 【修改说明】可选字段 secret 缺失时不生成 ERROR，符合字段规则设计。
        if not required:
            continue

        # 【修改说明】只有必要字段不存在时才记录 ERROR。
        if field_name not in sheet1_columns:
            # 【修改说明】记录具体缺失的必要字段，帮助用户定位 Excel 结构问题。
            add_issue(
                context,
                "ERROR",
                f"{SHEET1_NAME} 缺少必要字段：{field_name}",
            )

    # 【修改说明】检查 Sheet2 是否存在设备关联所需的 host 字段。
    if SHEET2_KEY_FIELD not in sheet2_columns:
        # 【修改说明】记录 Sheet2 缺少 host 的结构错误。
        add_issue(
            context,
            "ERROR",
            f"{SHEET2_NAME} 缺少必要字段：{SHEET2_KEY_FIELD}",
        )

def check_empty_headers(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    context: ValidationContext,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> None:
    """检查 Excel 是否存在空列名。"""
    for sheet_name, dataframe in (
        (SHEET1_NAME, context.sheet1),
        (SHEET2_NAME, context.sheet2),
    ):
        for index, column in enumerate(dataframe.columns, start=1):
            if not str(column).strip():
                add_issue(
                    context,
                    "ERROR",
                    f"{sheet_name} 第 {index} 列的列名为空。",
                )


def check_sheet1_data(
    context: ValidationContext,
) -> None:
    """按照字段规则检查 Sheet1。"""
# 【逐行说明】遍历 Sheet1 的每一行数据。
    for row_number, row in context.sheet1.iterrows():
# 【逐行说明】pandas 行索引从 0 开始，因此加 2 对应 Excel 实际数据行号。
        excel_row = row_number + 2

# 【逐行说明】读取当前设备 host，并统一转换成去空格字符串。
        host = str(row.get("host", "")).strip()

# 【逐行说明】检查 host 是否为空。
        if not host:
# 【逐行说明】记录 host 为空的问题。
            add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
                context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
                "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
                f"{SHEET1_NAME} Excel 第 {excel_row} 行：host 为空。",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            )

# 【逐行说明】遍历所有 Sheet1 字段及其必要性规则。
        for field_name, required in SHEET1_FIELD_RULES.items():
# 【逐行说明】如果当前字段不是必要字段，则跳过空值检查。
            if not required:
# 【逐行说明】当前代码行执行本步骤的具体操作。
                continue

# 【逐行说明】读取当前字段值，并统一转换为去空格字符串。
            value = str(row.get(field_name, "")).strip()

# 【逐行说明】判断必要字段是否为空。
            if not value:
# 【逐行说明】记录必要字段为空的问题。
                add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
                    context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
                    "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
                    (
# 【逐行说明】添加 Sheet 名称和 Excel 行号。
                        f"{SHEET1_NAME} Excel 第 {excel_row} 行，"
# 【逐行说明】添加设备 host；如果 host 为空则显示占位符。
                        f"host={host or '<EMPTY>'}："
# 【逐行说明】指出具体为空的必要字段。
                        f"必要字段 {field_name} 为空。"
# 【逐行说明】当前代码行执行本步骤的具体操作。
                    ),
# 【逐行说明】当前代码行执行本步骤的具体操作。
                )

# 【逐行说明】只有 host 列存在时才检查重复 host。
    if "host" in context.sheet1.columns:
# 【逐行说明】获取所有 host，并统一去除空格。
        hosts = context.sheet1["host"].astype(str).str.strip()
# 【逐行说明】找出重复且非空的 host。
        duplicates = hosts[hosts.duplicated(keep=False) & hosts.ne("")].unique()

# 【逐行说明】遍历所有重复 host。
        for host in duplicates:
# 【逐行说明】记录重复 host 问题。
            add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
                context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
                "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
                f"{SHEET1_NAME} 存在重复 host：{host}",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            )


# 【逐行说明】定义 Sheet2 数据检查函数。
def check_sheet2_data(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    context: ValidationContext,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> None:
    """检查 Sheet2 的 host、重复项和配置变量空值。"""
    if SHEET2_KEY_FIELD not in context.sheet2.columns:
        return

    hosts = context.sheet2[SHEET2_KEY_FIELD].astype(str).str.strip()

    for row_number, row in context.sheet2.iterrows():
        excel_row = row_number + 2
        host = str(row.get(SHEET2_KEY_FIELD, "")).strip()

        if not host:
            add_issue(
                context,
                "ERROR",
                f"{SHEET2_NAME} Excel 第 {excel_row} 行：host 为空。",
            )

        # Sheet2 除 host 外的字段均属于配置变量。
        for column in context.sheet2.columns:
            if column == SHEET2_KEY_FIELD:
                continue

            value = str(row.get(column, "")).strip()

            if not value:
                add_issue(
                    context,
                    "ERROR",
                    (
                        f"{SHEET2_NAME} Excel 第 {excel_row} 行，"
                        f"host={host or '<EMPTY>'}："
                        f"配置变量 {column} 为空。"
                    ),
                )

    duplicates = hosts[hosts.duplicated(keep=False) & hosts.ne("")].unique()

    for host in duplicates:
        add_issue(
            context,
            "ERROR",
            f"{SHEET2_NAME} 存在重复 host：{host}",
        )


def check_host_consistency(
    context: ValidationContext,
) -> None:
    """检查 Sheet1 和 Sheet2 的 host 是否完全一致。"""
# 【逐行说明】如果 Sheet1 没有 host，则无需重复检查。
    if "host" not in context.sheet1.columns:
# 【逐行说明】当前代码行执行本步骤的具体操作。
        return

# 【逐行说明】如果 Sheet2 没有 host，则同样交给结构检查处理。
    if SHEET2_KEY_FIELD not in context.sheet2.columns:
# 【逐行说明】当前代码行执行本步骤的具体操作。
        return

# 【逐行说明】构造 Sheet1 的非空 host 集合。
    sheet1_hosts = {
# 【逐行说明】遍历 Sheet1 中所有 host。
        value.strip()
# 【逐行说明】去掉 host 两端空格。
        for value in context.sheet1["host"].astype(str)
# 【逐行说明】只保留非空 host。
        if value.strip()
# 【逐行说明】当前代码行执行本步骤的具体操作。
    }

# 【逐行说明】构造 Sheet2 的非空 host 集合。
    sheet2_hosts = {
# 【逐行说明】遍历 Sheet2 中所有 host。
        value.strip()
# 【逐行说明】去掉 host 两端空格。
        for value in context.sheet2[SHEET2_KEY_FIELD].astype(str)
# 【逐行说明】只保留非空 host。
        if value.strip()
# 【逐行说明】当前代码行执行本步骤的具体操作。
    }

# 【逐行说明】计算只存在于 Sheet1 的 host。
    only_in_sheet1 = sheet1_hosts - sheet2_hosts
# 【逐行说明】计算只存在于 Sheet2 的 host。
    only_in_sheet2 = sheet2_hosts - sheet1_hosts

# 【逐行说明】按字典序遍历 Sheet1 独有的 host。
    for host in sorted(only_in_sheet1):
# 【逐行说明】记录 Sheet1 有而 Sheet2 没有的问题。
        add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
            context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            f"host={host} 存在于 Sheet1，但不存在于 Sheet2。",
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )

# 【逐行说明】按字典序遍历 Sheet2 独有的 host。
    for host in sorted(only_in_sheet2):
# 【逐行说明】记录 Sheet2 有而 Sheet1 没有的问题。
        add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
            context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            f"host={host} 存在于 Sheet2，但不存在于 Sheet1。",
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )


# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# PRE-CHECK：Jinja2 模板
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【逐行说明】定义模板读取和语法检查函数。
def load_template(context: ValidationContext) -> None:
    """读取模板并检查 Jinja2 语法。"""
    try:
        context.template_text = context.template_path.read_text(
            encoding="utf-8-sig"
        )
    except UnicodeDecodeError as exc:
        add_issue(
            context,
            "ERROR",
            f"配置模板不是有效的 UTF-8 文本文件：{exc}",
        )
        return
    except OSError as exc:
        add_issue(
            context,
            "ERROR",
            f"读取配置模板失败：{exc}",
        )
        return

    environment = Environment(undefined=StrictUndefined)

    try:
        parsed_template = environment.parse(context.template_text)
    except TemplateSyntaxError as exc:
        add_issue(
            context,
            "ERROR",
            (
                "Jinja2 模板语法错误："
                f"line={exc.lineno}, message={exc.message}"
            ),
        )
        return

    context.template_variables = meta.find_undeclared_variables(
        parsed_template
    )


def check_template_variables(
    context: ValidationContext,
) -> None:
    """检查模板变量是否全部存在于 Sheet2。"""
# 【逐行说明】获取 Sheet2 全部列名集合。
    sheet2_columns = set(context.sheet2.columns)

# 【逐行说明】说明 host 也是可被模板直接使用的变量。
    # host 是设备关联键，也允许模板直接使用 {{ host }}。
# 【逐行说明】当前可用变量集合直接采用 Sheet2 的所有列。
    available_variables = sheet2_columns

# 【逐行说明】计算模板需要但 Sheet2 没有提供的变量。
    missing_variables = context.template_variables - available_variables

# 【逐行说明】按变量名排序遍历缺失变量。
    for variable in sorted(missing_variables):
# 【逐行说明】记录模板变量缺失问题。
        add_issue(
# 【逐行说明】当前代码行执行本步骤的具体操作。
            context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            "ERROR",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            (
# 【逐行说明】当前代码行执行本步骤的具体操作。
                f"模板变量 {variable!r} 在 {SHEET2_NAME} 中不存在。"
# 【逐行说明】当前代码行执行本步骤的具体操作。
            ),
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )

# 【逐行说明】计算 Sheet2 中存在、但不是 host 且没有被模板使用的变量。
    unused_variables = (
# 【逐行说明】当前代码行执行本步骤的具体操作。
        sheet2_columns
# 【逐行说明】当前代码行执行本步骤的具体操作。
        - {SHEET2_KEY_FIELD}
# 【逐行说明】当前代码行执行本步骤的具体操作。
        - context.template_variables
# 【逐行说明】当前代码行执行本步骤的具体操作。
    )

# 【逐行说明】按变量名排序遍历多余变量。
    for variable in sorted(unused_variables):
# 【逐行说明】记录未使用变量。
        add_issue(
# 【逐行说明】设置 WARNING 级别，不阻止后续执行。
            context,
# 【逐行说明】当前代码行执行本步骤的具体操作。
            "WARNING",
# 【逐行说明】当前代码行执行本步骤的具体操作。
            (
# 【逐行说明】当前代码行执行本步骤的具体操作。
                f"{SHEET2_NAME} 中的变量 {variable!r} "
# 【逐行说明】当前代码行执行本步骤的具体操作。
                "没有被当前模板使用。"
# 【逐行说明】当前代码行执行本步骤的具体操作。
            ),
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )


# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# PRE-CHECK：Jinja2 Render
# 【逐行说明】原脚本注释用于说明代码结构或设计意图。
# ============================================================================


# 【逐行说明】定义为所有设备渲染最终配置的函数。
def render_devices(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    context: ValidationContext,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> None:
    """为每台设备渲染最终配置。"""
    if any(issue.level == "ERROR" for issue in context.issues):
        return

    environment = Environment(
        undefined=StrictUndefined,
        keep_trailing_newline=True,
    )

    try:
        template = environment.from_string(context.template_text)
    except TemplateSyntaxError as exc:
        add_issue(
            context,
            "ERROR",
            f"Jinja2 模板解析失败：{exc}",
        )
        return

    sheet1_by_host = {
        str(row["host"]).strip(): row.to_dict()
        for _, row in context.sheet1.iterrows()
    }

    for _, row in context.sheet2.iterrows():
        host = str(row[SHEET2_KEY_FIELD]).strip()
        variables = row.to_dict()

        connection_row = sheet1_by_host.get(host)

        if connection_row is None:
            # 理论上 Stage 4 已经发现，此处只是防御性检查。
            add_issue(
                context,
                "ERROR",
                f"渲染时找不到 host={host} 对应的 Sheet1 数据。",
            )
            continue

        # 【修改说明】将 Excel 中的设备管理地址映射到 Netmiko 的 host 参数，避免把设备名称误当成连接地址。
        connection_params = {
            "device_type": connection_row["device_type"],
            "host": connection_row["ip"],
            "username": connection_row["username"],
            "password": connection_row["password"],
# 【修改说明】将 Excel 中明确配置的 SSH/Telnet TCP 端口传递给 Netmiko，支持设备使用非默认端口。
            "port": int(connection_row["port"]),
        }

        secret = connection_row.get("secret", "")
        if secret:
            connection_params["secret"] = secret

        try:
            rendered_config = template.render(**variables)
        except Exception as exc:
            add_issue(
                context,
                "ERROR",
                f"host={host} Jinja2 渲染失败：{exc}",
            )
            continue

        if UNRENDERED_TEMPLATE_PATTERN.search(rendered_config):
            add_issue(
                context,
                "ERROR",
                (
                    f"host={host} 渲染结果中仍然存在未替换的 "
                    "Jinja2 模板标记。"
                ),
            )
            continue

        if not rendered_config.strip():
            add_issue(
                context,
                "ERROR",
                f"host={host} 渲染结果为空。",
            )
            continue

        context.devices.append(
            Device(
                host=host,
                connection_params=connection_params,
                variables=variables,
                rendered_config=rendered_config,
            )
        )


# ============================================================================
# PRE-CHECK：最终检查
# ============================================================================


def run_precheck(
    excel_path: Path,
    template_path: Path,
) -> ValidationContext:
    """执行完整 PRE-CHECK，整个过程不会连接网络设备。"""
    # 【修改说明】先创建空的 PRE-CHECK 上下文，使 Excel 结构读取失败时也能统一记录业务错误。
    context = ValidationContext(
        excel_path=excel_path,
        template_path=template_path,
        sheet1=pd.DataFrame(),
        sheet2=pd.DataFrame(),
        template_text="",
    )

    # 【修改说明】读取 Excel 的 Sheet1 和 Sheet2，并在读取失败时将异常转换为 PRE-CHECK ERROR。
    try:
        sheet1, sheet2 = load_excel(excel_path)
    except Exception as exc:
        # 【修改说明】将 Excel 打开、Sheet 缺失或读取失败统一记录为 PRE-CHECK ERROR，避免错误只出现在 CMD。
        add_issue(
            context,
            "ERROR",
            f"数据处理失败：{exc}",
        )
        # 【修改说明】返回包含错误信息的上下文，使主流程仍然能够生成完整 PRE-CHECK 报告和 precheck.log。
        return context

    # 【修改说明】保存清洗后的 Sheet1，统一处理空值和字符串空格。
    context.sheet1 = clean_dataframe(sheet1)
    # 【修改说明】保存清洗后的 Sheet2，统一处理空值和字符串空格。
    context.sheet2 = clean_dataframe(sheet2)

    # 【修改说明】检查两个 Sheet 是否存在空列名。
    check_empty_headers(context)
    # 【修改说明】检查 Sheet1/Sheet2 的结构，并确保可选字段 secret 不会被误判为缺失错误。
    check_required_columns(context)

    # 【修改说明】必要结构检查存在 ERROR 时，不继续执行依赖这些结构的后续检查，避免产生级联错误。
    if any(issue.level == "ERROR" for issue in context.issues):
        # 【修改说明】直接返回当前上下文，由主流程记录所有 PRE-CHECK ERROR 到 precheck.log。
        return context

    # 【修改说明】执行 Sheet1 数据检查，包括必要字段非空和重复 host 检查。
    check_sheet1_data(context)
    # 【修改说明】执行 Sheet2 数据检查，包括 host 非空、变量非空和重复 host 检查。
    check_sheet2_data(context)
    # 【修改说明】检查 Sheet1 与 Sheet2 的 host 是否完全一致。
    check_host_consistency(context)

    # 【修改说明】Excel 数据存在 ERROR 时不再继续解析模板，避免基于错误数据产生二次错误。
    if any(issue.level == "ERROR" for issue in context.issues):
        # 【修改说明】返回当前上下文，由主流程统一输出 PRE-CHECK 结果。
        return context

    # 【修改说明】读取配置模板并检查 Jinja2 语法。
    load_template(context)
    # 【修改说明】模板解析存在 ERROR 时停止后续变量检查和渲染，避免继续产生级联错误。
    if any(issue.level == "ERROR" for issue in context.issues):
        # 【修改说明】返回模板阶段的错误结果。
        return context

    # 【修改说明】检查模板需要的变量是否全部存在于 Sheet2。
    check_template_variables(context)
    # 【修改说明】模板变量存在缺失时禁止继续渲染设备配置。
    if any(issue.level == "ERROR" for issue in context.issues):
        # 【修改说明】返回变量检查失败的结果。
        return context

    # 【修改说明】对所有设备执行 Jinja2 渲染，确保真正下发前已经得到完整配置。
    render_devices(context)

    # 【修改说明】返回完整 PRE-CHECK 上下文，供报告和后续 Preview/Deploy 使用。
    return context

def print_precheck_report(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    context: ValidationContext,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    logger: logging.Logger,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> tuple[int, int]:
    """输出 PRE-CHECK 报告，并返回 ERROR 和 WARNING 数量。"""
    print()
    print("=" * 72)
    print("PRE-CHECK RESULT")
    print("=" * 72)

    error_count = 0
    warning_count = 0

    for issue in context.issues:
        logger.log(
            logging.ERROR if issue.level == "ERROR" else logging.WARNING,
            issue.message,
        )

        if issue.level == "ERROR":
            error_count += 1
        elif issue.level == "WARNING":
            warning_count += 1

    print()
    print(f"ERROR   : {error_count}")
    print(f"WARNING : {warning_count}")

    if error_count:
        print()
        print("[FAILED] PRE-CHECK 未通过。")
        print("不会连接任何网络设备。")
        return error_count, warning_count

    print()
    print(f"[PASSED] PRE-CHECK 通过。")
    print(f"设备数量：{len(context.devices)}")
# 【修改说明】返回统计结果，使主流程可以只在存在 ERROR/WARNING 时创建 precheck.log。
    return error_count, warning_count


def write_precheck_log(
# 【修改说明】接收本次 PRE-CHECK 的执行目录和检查结果，用于按需生成阶段日志。
    execution_dir: Path,
# 【修改说明】接收 PRE-CHECK 中已经收集到的 ERROR/WARNING，避免重新执行检查。
    context: ValidationContext,
) -> None:
    """仅在 PRE-CHECK 存在 ERROR/WARNING 时创建并写入 precheck.log。"""
# 【修改说明】只筛选真正需要落盘的 ERROR 和 WARNING，INFO 不触发文件创建。
    issues_to_log = [
        issue
        for issue in context.issues
        if issue.level in {"ERROR", "WARNING"}
    ]

# 【修改说明】没有 ERROR/WARNING 时直接返回，因此正常 PRE-CHECK 不会创建 precheck.log。
    if not issues_to_log:
        return

# 【修改说明】创建只写入文件的 PRE-CHECK Logger，避免已经在 CMD 显示过的问题再次打印到终端。
    precheck_logger = setup_logger(
        execution_dir / "precheck.log",
        console_output=False,
    )

# 【修改说明】按原始检查顺序重新写入所有 ERROR/WARNING，确保日志文件完整保留异常和告警信息。
    for issue in issues_to_log:
# 【修改说明】根据 ValidationIssue 的级别写入对应的 logging 级别。
        precheck_logger.log(
            logging.ERROR if issue.level == "ERROR" else logging.WARNING,
            issue.message,
        )


# ============================================================================
# 配置文件生成
# ============================================================================


def create_output_directory() -> Path:
    """创建本次脚本执行的独立日志目录，不提前创建 Preview 子目录。"""
    # 【修改说明】使用精确到秒的时间戳作为本次执行目录名称，目录名称直观反映执行时间。
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # 【修改说明】将本次脚本执行产生的全部文件统一放到脚本目录下的 logs 根目录。
    logs_dir = SCRIPT_DIR / "logs"
    # 【修改说明】根据时间戳构造本次执行的独立目录路径。
    execution_dir = logs_dir / timestamp
    # 【修改说明】创建执行目录；父级 logs 不存在时一并创建，但不创建 preview 子目录。
    try:
        execution_dir.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        # 【修改说明】同一秒内重复执行时追加递增序号，避免覆盖之前的执行结果。
        sequence = 1
        # 【修改说明】为同秒重复执行构造带序号的候选目录名称。
        candidate_dir = logs_dir / f"{timestamp}_{sequence:02d}"
        # 【修改说明】持续寻找尚未存在的目录，保证每次执行都有独立保存位置。
        while candidate_dir.exists():
            # 【修改说明】递增序号以处理同一秒内已经存在多个执行目录的情况。
            sequence += 1
            # 【修改说明】重新生成下一组候选目录名称。
            candidate_dir = logs_dir / f"{timestamp}_{sequence:02d}"
        # 【修改说明】将最终未占用的候选目录作为本次执行目录。
        execution_dir = candidate_dir
        # 【修改说明】创建最终确定的本次执行目录。
        execution_dir.mkdir(parents=True, exist_ok=False)

    # 【修改说明】返回仅包含本次执行根目录的路径，preview 将在真正写配置文件时惰性创建。
    return execution_dir

def write_config_files(
    devices: list[Device],
    output_dir: Path,
) -> None:
    """为每台设备生成最终配置文件，并在首次写入时惰性创建 Preview 目录。"""
    # 【修改说明】只有已经通过 PRE-CHECK 且拥有可写入的渲染结果时，才创建 preview 子目录。
    preview_dir = output_dir / "preview"
    # 【修改说明】在真正开始生成配置文件之前创建 Preview 目录，避免 PRE-CHECK 失败产生空目录。
    preview_dir.mkdir(parents=False, exist_ok=False)

    # 【修改说明】遍历所有已经成功完成 Jinja2 渲染的设备。
    for device in devices:
        # 【修改说明】为当前设备构造独立的最终配置文件路径。
        config_file = preview_dir / f"{device.host}.txt"
        # 【修改说明】使用 UTF-8 将最终渲染配置完整写入设备配置文件。
        config_file.write_text(
            device.rendered_config,
            encoding="utf-8",
        )

def normalize_config_commands(rendered_config: str) -> list[str]:
    """把最终配置文本转换为 Netmiko 可接受的命令列表。"""
    commands: list[str] = []

    for line in rendered_config.splitlines():
        command = line.rstrip()

        if not command.strip():
            continue

        commands.append(command)

    return commands


def find_config_errors(output: str) -> list[str]:
    """根据常见 CLI 错误关键字检查配置执行结果。"""
# 【逐行说明】创建错误行列表。
    errors: list[str] = []

# 【逐行说明】按行遍历设备 CLI 输出。
    for line in output.splitlines():
# 【逐行说明】去除当前输出行两端空白。
        stripped_line = line.strip()

# 【逐行说明】忽略空输出行。
        if not stripped_line:
# 【逐行说明】当前代码行执行本步骤的具体操作。
            continue

# 【逐行说明】判断当前行是否匹配任意错误正则。
        if any(
# 【逐行说明】对当前行执行不区分大小写的正则匹配。
            re.search(pattern, stripped_line, re.IGNORECASE)
# 【逐行说明】遍历所有预定义 CLI 错误模式。
            for pattern in CONFIG_ERROR_PATTERNS
# 【逐行说明】当前代码行执行本步骤的具体操作。
        ):
# 【逐行说明】如果命中错误模式，则保存原始错误行。
            errors.append(stripped_line)

# 【逐行说明】返回检测到的错误行列表。
    return errors


# 【逐行说明】定义设备配置保存函数。
def save_device_config(connection: Any) -> str:
    """
    保存设备配置。

    优先使用 Netmiko 当前平台实现的 save_config()。
    Netmiko 会根据 device_type 使用对应平台的保存机制。
    """
# 【逐行说明】调用 Netmiko 平台对应的 save_config() 保存运行配置。
    output = connection.save_config()

# 【逐行说明】检查保存操作返回的 CLI 输出中是否存在错误。
    errors = find_config_errors(output)

# 【逐行说明】如果发现保存错误，则抛出异常。
    if errors:
# 【逐行说明】创建保存失败异常。
        raise RuntimeError(
# 【逐行说明】添加固定的错误说明。
            "保存配置时设备返回错误："
# 【逐行说明】将所有检测到的错误行拼接到异常消息中。
            + " | ".join(errors)
# 【逐行说明】当前代码行执行本步骤的具体操作。
        )

# 【逐行说明】返回设备保存命令的原始输出。
    return output


# 【逐行说明】定义单台设备完整 Deploy 生命周期函数。
def deploy_one_device(
# 【逐行说明】当前代码行执行本步骤的具体操作。
    device: Device,
# 【逐行说明】当前代码行执行本步骤的具体操作。
    output_dir: Path,
# 【逐行说明】当前代码行执行本步骤的具体操作。
) -> DeploymentResult:
    """执行单台设备的完整 Deploy 生命周期。"""
    connection = None
# 【修改说明】记录实际生成的 Preview 配置文件路径，便于 Deploy 汇总准确指向设备配置文件。
    config_file = output_dir / "preview" / f"{device.host}.txt"

    try:
        connection = ConnectHandler(
            **device.connection_params,
        )

        commands = normalize_config_commands(
            device.rendered_config,
        )

        if not commands:
            return DeploymentResult(
                host=device.host,
                status="FAILED",
                message="最终配置为空。",
                config_file=str(config_file),
            )

        output = connection.send_config_set(
            config_commands=commands,
            enter_config_mode=True,
            exit_config_mode=True,
        )

        config_errors = find_config_errors(output)

        if config_errors:
            return DeploymentResult(
                host=device.host,
                status="FAILED",
                message=(
                    "配置命令执行失败："
                    + " | ".join(config_errors)
                ),
                config_file=str(config_file),
            )

        try:
            save_device_config(connection)
        except Exception as exc:
            return DeploymentResult(
                host=device.host,
                status="PARTIAL_SUCCESS",
                message=f"配置下发成功，但保存失败：{exc}",
                config_file=str(config_file),
            )

        return DeploymentResult(
            host=device.host,
            status="SUCCESS",
            message="配置下发并保存成功。",
            config_file=str(config_file),
        )

    except AuthenticationException as exc:
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=(
                "认证失败：用户名或密码错误的可能性最高。\n"
                "可能原因：SSH 密钥、AAA 策略、账号状态或认证方式导致设备拒绝登录。\n"
                f"Netmiko 原始信息：{exc}"
            ),
            config_file=str(config_file),
        )

    except SSHException as exc:
        if "Error reading SSH protocol banner" in str(exc):
            return DeploymentResult(
                host=device.host,
                status="FAILED",
                message=(
                    "SSH 协议握手失败：无法读取 SSH protocol banner。\n"
                    "可能原因：IP 或端口对应的服务不是 SSH、SSH 服务未正常监听。\n"
                    "也可能是防火墙或 ACL 允许 TCP 建连但阻断后续 SSH 握手，"
                    "或者设备响应过慢。\n"
                    "请重点检查 Excel 中的 ip 和 port，并确认该端口确实提供 SSH 服务。\n"
                    f"Paramiko 原始信息：{exc}"
                ),
                config_file=str(config_file),
            )
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=f"SSH 连接失败：{exc}",
            config_file=str(config_file),
        )

    except NetmikoTimeoutException as exc:
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=(
                "连接超时：可能是 IP 地址不可达、TCP 端口不可达、"
                "防火墙/ACL 丢弃连接，或者设备响应超时。"
                f" Netmiko 原始信息：{exc}"
            ),
            config_file=str(config_file),
        )

    except ConnectionException as exc:
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=f"设备连接失败：{exc}",
            config_file=str(config_file),
        )

    except ConfigInvalidException as exc:
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=f"进入配置模式失败：{exc}",
            config_file=str(config_file),
        )

    except Exception as exc:
        return DeploymentResult(
            host=device.host,
            status="FAILED",
            message=f"未知异常：{type(exc).__name__}: {exc}",
            config_file=str(config_file),
        )

    finally:
        if connection is not None:
            try:
                connection.disconnect()
            except Exception:
                # disconnect 本身失败不能覆盖前面的真实执行结果。
                pass


def deploy_devices(
    devices: list[Device],
    output_dir: Path,
    logger: logging.Logger,
) -> list[DeploymentResult]:
    """使用线程池并发执行多台设备配置下发。"""
# 【逐行说明】初始化最终结果列表。
    results: list[DeploymentResult] = []

# 【逐行说明】线程数不超过 MAX_WORKERS，也不超过设备数量；空设备列表至少使用 1。
    worker_count = min(MAX_WORKERS, max(len(devices), 1))

# 【逐行说明】创建固定大小的线程池，并在代码块结束时等待所有任务完成。
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
# 【逐行说明】创建 Future 到设备 host 的映射，便于异步结果返回时识别设备。
        future_map = {
# 【逐行说明】向线程池提交单台设备 Deploy 任务。
            executor.submit(
# 【逐行说明】指定执行函数。
                deploy_one_device,
# 【逐行说明】传入当前设备。
                device,
# 【逐行说明】传入输出目录。
                output_dir,
# 【逐行说明】用设备 host 作为 Future 的标识值。
            ): device.host
# 【逐行说明】当前代码行执行本步骤的具体操作。
            for device in devices
# 【逐行说明】当前代码行执行本步骤的具体操作。
        }

# 【逐行说明】按完成先后顺序取得 Future，而不是按提交顺序等待。
        for future in as_completed(future_map):
# 【逐行说明】根据 Future 找到对应设备 host。
            host = future_map[future]

# 【逐行说明】捕获线程任务本身抛出的异常。
            try:
# 【逐行说明】获取线程执行结果。
                result = future.result()
# 【逐行说明】如果线程函数意外抛异常，则构造统一 FAILED 结果。
            except Exception as exc:
# 【逐行说明】创建失败结果对象。
                result = DeploymentResult(
# 【逐行说明】设置设备 host。
                    host=host,
# 【逐行说明】设置 FAILED 状态。
                    status="FAILED",
# 【逐行说明】记录线程执行异常信息。
                    message=f"线程执行异常：{exc}",
# 【逐行说明】当前代码行执行本步骤的具体操作。
                )

# 【逐行说明】将当前设备结果加入最终结果列表。
            results.append(result)

# 【修改说明】根据设备最终状态选择对应的 logging 级别，使 SUCCESS、PARTIAL_SUCCESS 和 FAILED 在 CMD 与文件中有一致的严重程度。
            if result.status == "SUCCESS":
                # 【修改说明】成功结果使用 INFO 级别记录，因为设备已经完成配置并保存。
                log_level = logging.INFO
            elif result.status == "PARTIAL_SUCCESS":
                # 【修改说明】部分成功使用 WARNING 级别记录，因为配置已经执行但后续保存等步骤存在问题。
                log_level = logging.WARNING
            else:
                # 【修改说明】失败结果使用 ERROR 级别记录，方便用户通过日志级别快速定位失败设备。
                log_level = logging.ERROR

            # 【修改说明】将单台设备的标题、状态和完整异常信息作为一次日志记录输出，避免多线程执行时不同设备的错误信息互相穿插。
            # 【修改说明】在每台设备标题前增加两个空行，并使用 raw 日志记录，使设备标题不显示日期、时间和日志级别。
            logger.log(
                log_level,
                "\n\n========== 设备：%s ==========",
                result.host,
                extra={"raw": True},
            )
            # 【修改说明】使用普通日志记录设备执行结果，使状态、时间和日志级别仍然保留在详细执行记录中。
            logger.log(
                log_level,
                "[%s] %s",
                result.status,
                result.message,
            )

# 【逐行说明】按设备 host 排序后返回结果，使最终输出顺序稳定。
    return sorted(results, key=lambda item: item.host)


# 【逐行说明】定义 Deploy 最终汇总文件生成函数。
def log_deployment_summary(
    results: list[DeploymentResult],
    logger: logging.Logger,
) -> None:
    """将 Deploy 最终汇总追加到 deployment.log。"""
    # 【修改说明】统计完全成功的设备数量，供 Deployment Summary 使用。
    success_count = sum(
        result.status == "SUCCESS"
        for result in results
    )
    # 【修改说明】统计配置成功但保存等后续步骤存在问题的设备数量。
    partial_count = sum(
        result.status == "PARTIAL_SUCCESS"
        for result in results
    )
    # 【修改说明】统计执行失败的设备数量。
    failed_count = sum(
        result.status == "FAILED"
        for result in results
    )

    # 【修改说明】在 Summary 标题前增加两个空行，并使用 raw 日志记录，使标题保持纯文本显示。
    logger.info("\n\n========== Deployment Summary ==========", extra={"raw": True})
    # 【修改说明】记录 SUCCESS 数量。
    logger.info("SUCCESS         : %s", success_count)
    # 【修改说明】记录 PARTIAL_SUCCESS 数量。
    logger.info("PARTIAL_SUCCESS : %s", partial_count)
    # 【修改说明】记录 FAILED 数量。
    logger.info("FAILED          : %s", failed_count)
    # 【修改说明】增加空白分隔，使最终统计区域与前面的设备执行日志明显区分。
    logger.info("")
    # 【修改说明】标记设备最终状态统计开始，避免在 Summary 中重复输出完整异常堆栈。
    logger.info("Device Results")

    # 【修改说明】按照 host 排序输出设备最终状态，使 Summary 中的设备顺序稳定。
    for result in results:
        # 【修改说明】只记录 host 和最终状态，让 Summary 保持简洁，详细错误已经在前面的设备日志块中记录。
        logger.info(
            "%s | %s",
            result.host,
            result.status,
        )

def main() -> None:
    """程序主入口。"""
    # 【修改说明】输出程序标题分隔线，帮助用户快速识别工具启动位置。
    print("=" * 72)
    # 【修改说明】输出程序名称。
    print("Network Devices Config Tool")
    # 【修改说明】输出标题结束分隔线。
    print("=" * 72)
    # 【修改说明】输出空行，提升 CLI 可读性。
    print()

    # 【修改说明】初始化本次执行目录变量，便于异常处理阶段判断日志是否已经建立。
    execution_dir: Path | None = None
    # 【修改说明】初始化 PRE-CHECK Logger，避免在日志目录创建之前无法记录异常。
    logger: logging.Logger | None = None

    # 【修改说明】捕获主流程中的可预期和非预期异常，并在 finally 中统一暂停窗口。
    try:
        # 【修改说明】让用户选择 Excel 和模板文件，并取得完整路径。
        excel_path, template_path = prompt_for_files()

        # 【修改说明】输出空行，分隔文件输入和后续模式选择。
        print()
        # 【修改说明】显示最终选中的 Excel 文件名。
        print(f"[OK] Excel    : {excel_path.name}")
        # 【修改说明】显示最终选中的模板文件名。
        print(f"[OK] Template : {template_path.name}")

        # 【修改说明】让用户选择 Preview 或 Deploy。
        mode = prompt_for_mode()

        # 【修改说明】创建本次执行的唯一工作目录；从这里开始本次执行产生的日志集中保存。
        execution_dir = create_output_directory()
        # 【修改说明】初始化仅输出到 CMD 的 Logger，避免 PRE-CHECK 正常时提前创建日志文件。
        logger = setup_logger()

        # 【修改说明】执行完整 PRE-CHECK；此阶段仍然不会调用 ConnectHandler。
        context = run_precheck(
            excel_path=excel_path,
            template_path=template_path,
        )

        # 【修改说明】输出 PRE-CHECK 报告；此时 ERROR/WARNING 仍然会实时显示在 CMD。
        error_count, warning_count = print_precheck_report(
            context=context,
            logger=logger,
        )

        # 【修改说明】只有 PRE-CHECK 存在 ERROR 或 WARNING 时才创建 precheck.log，正常检查不会留下该文件。
        write_precheck_log(
            execution_dir=execution_dir,
            context=context,
        )

        # 【修改说明】判断 PRE-CHECK 是否存在 ERROR；只要存在一个 ERROR 就禁止进入设备连接阶段。
        if any(issue.level == "ERROR" for issue in context.issues):
            # 【修改说明】输出 PRE-CHECK 失败后的安全提示，明确说明不会连接设备。
            print("\n[FAILED] PRE-CHECK 未通过，不会连接任何网络设备。")
            # 【修改说明】输出本次执行目录，方便用户直接查看 precheck.log。
            print(f"本次执行目录：{execution_dir}")
            # 【修改说明】结束本次执行，finally 会统一等待用户按 Enter 后再退出 CMD。
            return

        # 【修改说明】PRE-CHECK 全部通过后，才生成本次执行的设备配置文件。
        write_config_files(
            devices=context.devices,
            output_dir=execution_dir,
        )

        # 【修改说明】输出配置生成完成的提示。
        print()
        # 【修改说明】输出配置生成阶段分隔线。
        print("=" * 72)
        # 【修改说明】明确说明配置文件已经生成。
        print("配置文件生成完成")
        # 【修改说明】结束配置生成阶段分隔线。
        print("=" * 72)
        # 【修改说明】输出本次执行目录，便于定位全部日志和配置文件。
        print(f"本次执行目录：{execution_dir}")
        # 【修改说明】输出 preview 配置文件目录。
        print(f"配置文件目录：{execution_dir / 'preview'}")
        # 【修改说明】输出成功生成的配置文件数量。
        print(f"配置文件数量：{len(context.devices)}")

        # 【修改说明】Preview 模式只生成配置文件，不允许进入网络设备连接阶段。
        if mode == "preview":
            # 【修改说明】明确说明 Preview 不会连接任何网络设备。
            print("\n[PREVIEW] 不会连接任何网络设备。")
            # 【修改说明】明确说明 Preview 工作已经完成。
            print("[PREVIEW] 配置文件已经生成，程序结束。")
            # 【修改说明】结束 Preview 模式，finally 会统一暂停窗口。
            return

        # 【修改说明】要求用户严格输入大写 YES，作为真正接触生产设备前的二次确认。
        if not prompt_for_deploy_confirmation():
            # 【修改说明】明确说明用户未确认，因此不会连接任何设备。
            print("\n[CANCELLED] 用户未输入 YES，不会连接任何网络设备。")
            # 【修改说明】结束取消流程，finally 会统一暂停窗口。
            return

        # 【修改说明】输出 Deploy 开始分隔线。
        print()
        print("=" * 72)
        # 【修改说明】明确提示开始执行网络设备配置下发。
        print("开始 DEPLOY")
        # 【修改说明】结束 Deploy 标题分隔线。
        print("=" * 72)

        # 【修改说明】创建 Deployment 专用日志文件，并恢复 CMD 输出，使设备执行结果可以实时显示并同时保存。
        deployment_logger = setup_logger(execution_dir / "deployment.log")
        # 【修改说明】将主流程当前使用的 Logger 切换为 Deployment Logger，使后续 Deploy 异常写入 deployment.log。
        logger = deployment_logger

        # 【修改说明】启动多设备并发配置下发。
        results = deploy_devices(
            devices=context.devices,
            output_dir=execution_dir,
            logger=deployment_logger,
        )

        # 【修改说明】将最终统计和设备明细直接追加到 deployment.log，取消单独的 deployment_summary.txt。
        log_deployment_summary(
            results=results,
            logger=deployment_logger,
        )

        # 【修改说明】统计完全成功数量。
        success_count = sum(
            result.status == "SUCCESS"
            for result in results
        )
        # 【修改说明】统计部分成功数量。
        partial_count = sum(
            result.status == "PARTIAL_SUCCESS"
            for result in results
        )
        # 【修改说明】统计失败数量。
        failed_count = sum(
            result.status == "FAILED"
            for result in results
        )

        # 【修改说明】输出 Deploy 完成分隔线。
        print()
        print("=" * 72)
        # 【修改说明】输出 Deploy 完成提示。
        print("DEPLOY 完成")
        # 【修改说明】输出结果统计区域分隔线。
        print("=" * 72)
        # 【修改说明】输出 SUCCESS 数量。
        print(f"SUCCESS         : {success_count}")
        # 【修改说明】输出 PARTIAL_SUCCESS 数量。
        print(f"PARTIAL_SUCCESS : {partial_count}")
        # 【修改说明】输出 FAILED 数量。
        print(f"FAILED          : {failed_count}")
        # 【修改说明】输出任务结果目录。
        print(f"本次执行目录     : {execution_dir}")
        # 【修改说明】明确提示详细 Deploy 日志和最终汇总都位于 deployment.log。
        print(f"Deployment 日志  : {execution_dir / 'deployment.log'}")

    # 【修改说明】捕获用户按 Ctrl+C 主动中断程序，并保留 CMD 窗口供用户查看现场信息。
    except KeyboardInterrupt:
        # 【修改说明】输出用户中断警告。
        print("\n[WARNING] 用户中断程序。")
        # 【修改说明】如果 PRE-CHECK 日志已经建立，则将用户中断事件写入日志。
        if logger is not None:
            logger.warning("用户中断程序。")

    # 【修改说明】捕获文件不存在异常，并写入已经建立的阶段日志。
    except FileNotFoundError as exc:
        # 【修改说明】输出文件不存在错误。
        print(f"\n[ERROR] 文件不存在：{exc}")
        # 【修改说明】在 PRE-CHECK 阶段发生文件错误时创建 precheck.log，保证 ERROR 有文件记录。
        if execution_dir is not None and logger is not None:
            logger.error("文件不存在：%s", exc)
            if not (execution_dir / "deployment.log").exists():
                logger = setup_logger(execution_dir / "precheck.log")
                logger.error("文件不存在：%s", exc)

    # 【修改说明】捕获文件或目录权限不足异常，并写入已经建立的阶段日志。
    except PermissionError as exc:
        # 【修改说明】输出权限错误。
        print(f"\n[ERROR] 文件权限不足：{exc}")
        # 【修改说明】在 PRE-CHECK 阶段发生权限错误时创建 precheck.log，保证 ERROR 有文件记录。
        if execution_dir is not None and logger is not None:
            logger.error("文件权限不足：%s", exc)
            if not (execution_dir / "deployment.log").exists():
                logger = setup_logger(execution_dir / "precheck.log")
                logger.error("文件权限不足：%s", exc)

    # 【修改说明】捕获数据处理常见的 ValueError 和 KeyError，并写入阶段日志。
    except (ValueError, KeyError) as exc:
        # 【修改说明】输出数据处理错误。
        print(f"\n[ERROR] 数据处理失败：{exc}")
        # 【修改说明】在 PRE-CHECK 阶段发生数据错误时创建 precheck.log，保证 ERROR 有文件记录。
        if execution_dir is not None and logger is not None:
            logger.error("数据处理失败：%s", exc)
            if not (execution_dir / "deployment.log").exists():
                logger = setup_logger(execution_dir / "precheck.log")
                logger.error("数据处理失败：%s", exc)

    # 【修改说明】捕获其他未处理异常，避免程序静默退出，并把异常记录到阶段日志。
    except Exception as exc:
        # 【修改说明】输出错误类别和具体异常信息。
        print(
            "\n[ERROR] 程序发生未处理异常："
            f"{type(exc).__name__}: {exc}"
        )
        # 【修改说明】在 PRE-CHECK 阶段发生未处理异常时创建 precheck.log，保证 ERROR 有文件记录。
        if execution_dir is not None and logger is not None:
            if (execution_dir / "deployment.log").exists():
                logger.exception("程序发生未处理异常：%s", exc)
            else:
                logger = setup_logger(execution_dir / "precheck.log")
                logger.exception("程序发生未处理异常：%s", exc)

    # 【修改说明】无论程序从哪个 return 或异常分支离开，都统一暂停 CMD 窗口。
    finally:
        # 【修改说明】提示用户按 Enter 后退出，使 Preview、Deploy、错误和取消等所有路径都保留可查看的终端内容。
        try:
            input("\n请按 Enter 键退出...")
        except (EOFError, KeyboardInterrupt):
            # 【修改说明】非交互环境或用户再次中断等待时直接结束，不让退出等待本身产生新的异常。
            pass

if __name__ == "__main__":
# 【逐行说明】直接运行脚本时调用主入口函数。
    main()
