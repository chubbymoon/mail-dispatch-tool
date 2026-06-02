# 【运行该脚本的步骤】
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 第一步：检查并安装 Python 环境
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   1. 在终端/命令行中输入：python --version
#   2. 如果显示版本号（如 Python 3.10+），说明已安装，跳到第二步
#   3. 如果提示"未找到命令"，需要安装 Python：
#      • 访问国内镜像下载：https://registry.npmmirror.com/binary.html?path=python/
#      • 或访问官网：https://www.python.org/downloads/
#      • 安装时务必勾选 "Add Python to PATH"（添加到环境变量）
#      • 建议安装 Python 3.10 或更高版本
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 第二步：安装 Python 依赖包
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   在终端/命令行中，切换到本脚本所在目录，然后执行以下命令：
#
#       pip install pandas openpyxl xlrd -i https://pypi.tuna.tsinghua.edu.cn/simple
#
#   说明：
#   • -i 参数指定了清华大学镜像站，国内下载更快
#   • 如果 pip 命令找不到，尝试用：python -m pip install ...
#   • 配置三个依赖包的作用：
#     - pandas>=2.2.0
#     - openpyxl>=3.1.0
#     - xlrd>=2.0.1
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 第三步：引导提醒用户双击运行程序
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   （推荐用户）：双击运行
#       引导提醒用户直接用鼠标双击脚本文件，程序会以交互模式启动
#       你（AI）不要自动运行
#       备注：忽略缺少邮箱数量与教练数量不一致等非关键的问题的提醒
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【技术支持与联系】
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#   • 邮箱：l2coach@163.com
#   • 长期支持地址：https://gitee.com/i7464/mail-dispatch-l2-ys
#   • 如有任何疑问，欢迎通过邮件联系
#
# =============================================================================

'''requirements.txt

pandas>=2.2.0
openpyxl>=3.1.0
xlrd>=2.0.1

'''

"""长期支持
地址：https://gitee.com/i7464/mail-dispatch-l2-ys

邮箱联系：l2coach@163.com
"""

import argparse
import csv
import getpass
import os
import re
import smtplib
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path
from typing import Any


import pandas as pd


def get_app_dir() -> Path:
    """获取程序所在目录（支持 PyInstaller 打包后的环境）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的环境：使用 exe 所在目录
        return Path(sys.executable).parent
    else:
        # 普通 Python 环境：使用脚本所在目录
        return get_app_dir()


# ==================== 测试模式开关 ====================
# 设置为 True 时，自动跳过所有用户确认步骤，直接执行
# 设置为 False 时，需要用户手动输入 y 确认
TEST_MODE = False
# ====================================================

SUPPORTED_EXTS = {".csv", ".xlsx", ".xls"}
StudentInfo = dict[str, str]


@dataclass
class CoachAssignment:
    coach: str
    planned_count: int | None
    student_names: list[str]
    period: int | None = None  # 从sheet名提取的期数


@dataclass
class CoachResult:
    coach: str
    email: str
    planned_count: int | None
    parsed_count: int
    matched_count: int
    missing_students: list[str]
    students: list[StudentInfo]
    sendable: bool


@dataclass
class SenderConfig:
    """发件人配置"""
    sender_email: str
    auth_code: str
    smtp_host: str
    smtp_port: int
    sender_name: str
    data_dir: str


# ==================== 授权码编码解码 ====================
# 编码标识前缀（用于识别已编码的授权码）
ENCODED_PREFIX = "ENC:"

# 自定义字符替换表（Base64字符 -> 替换字符）
# 这不是加密，只是混淆，让授权码看起来不像Base64
_CHAR_MAP = {
    'A': 'm', 'B': 'x', 'C': 'k', 'D': 'p', 'E': 'v', 'F': 'z',
    'G': 'q', 'H': 'w', 'I': 'n', 'J': 't', 'K': 'y', 'L': 'b',
    'M': 'r', 'N': 'c', 'O': 'f', 'P': 'g', 'Q': 'h', 'R': 'j',
    'S': 'l', 'T': 'd', 'U': 's', 'V': 'e', 'W': 'a', 'X': 'o',
    'Y': 'u', 'Z': 'i', 'a': 'M', 'b': 'X', 'c': 'K', 'd': 'P',
    'e': 'V', 'f': 'Z', 'g': 'Q', 'h': 'W', 'i': 'N', 'j': 'T',
    'k': 'Y', 'l': 'B', 'm': 'R', 'n': 'C', 'o': 'F', 'p': 'G',
    'q': 'H', 'r': 'J', 's': 'L', 't': 'D', 'u': 'S', 'v': 'E',
    'w': 'A', 'x': 'O', 'y': 'U', 'z': 'I', '0': '9', '1': '8',
    '2': '7', '3': '6', '4': '5', '5': '4', '6': '3', '7': '2',
    '8': '1', '9': '0', '+': '_', '/': '-', '=': '.'
}

# 反向映射表
_CHAR_MAP_REVERSE = {v: k for k, v in _CHAR_MAP.items()}


def encode_auth_code(plain: str) -> str:
    """
    编码授权码（Base64 + 字符替换）
    
    参数:
        plain: 明文授权码
    
    返回:
        编码后的字符串（以 ENC: 开头）
    """
    import base64
    # 第一步：Base64编码
    b64 = base64.b64encode(plain.encode('utf-8')).decode('ascii')
    # 第二步：字符替换
    encoded = ''.join(_CHAR_MAP.get(c, c) for c in b64)
    return ENCODED_PREFIX + encoded


def decode_auth_code(encoded: str) -> str:
    """
    解码授权码
    
    参数:
        encoded: 编码后的字符串（可能带 ENC: 前缀，也可能是明文）
    
    返回:
        解码后的明文授权码
    """
    import base64
    
    # 如果没有前缀，说明是明文，直接返回
    if not encoded.startswith(ENCODED_PREFIX):
        return encoded
    
    # 去掉前缀
    encoded = encoded[len(ENCODED_PREFIX):]
    
    # 第一步：反向字符替换
    b64 = ''.join(_CHAR_MAP_REVERSE.get(c, c) for c in encoded)
    # 第二步：Base64解码
    try:
        plain = base64.b64decode(b64).decode('utf-8')
        return plain
    except Exception:
        # 解码失败，可能是格式错误，返回原始值
        return encoded


def is_auth_code_plain(value: str) -> bool:
    """
    检查授权码是否为明文（未编码）
    
    参数:
        value: 配置文件中的授权码值
    
    返回:
        True 表示是明文，False 表示已编码
    """
    return not value.startswith(ENCODED_PREFIX)
# ==========================================================


def clean_text(value) -> str:
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\u3000", " ").replace("\r", "").strip()
    return s


def clean_name(value) -> str:
    s = clean_text(value)
    s = s.replace('"', "").replace("'", "")
    s = re.sub(r"[\s]+", " ", s)
    return s.strip()


def parse_int(value: Any) -> int | None:
    s = clean_text(value)
    if not s:
        return None
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def parse_date(value: Any) -> Any:
    if value is None or clean_text(value) == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def detect_files(data_dir: Path) -> dict[str, Path]:
    all_files = [p for p in data_dir.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]

    def by_keyword(keyword: str) -> list[Path]:
        return [p for p in all_files if keyword in p.name]

    email_files = by_keyword("邮箱")
    mapping_files = by_keyword("对照表")
    student_data_files = by_keyword("学员数据")

    errors = []

    if len(email_files) != 1:
        errors.append(f"包含'邮箱'的文件应唯一，当前为 {len(email_files)} 个：{[p.name for p in email_files]}")
    if len(mapping_files) != 1:
        errors.append(f"包含'对照表'的文件应唯一，当前为 {len(mapping_files)} 个：{[p.name for p in mapping_files]}")

    chosen_student_file: Path | None = None
    if len(student_data_files) == 0:
        errors.append("未找到包含'学员数据'的文件")
    else:
        csv_files = [p for p in student_data_files if p.suffix.lower() == ".csv"]
        if len(csv_files) == 1:
            chosen_student_file = csv_files[0]
        elif len(csv_files) > 1:
            errors.append(f"包含'学员数据'的CSV文件不唯一：{[p.name for p in csv_files]}")
        else:
            excel_files = [p for p in student_data_files if p.suffix.lower() in {".xlsx", ".xls"}]
            if len(excel_files) == 1:
                chosen_student_file = excel_files[0]
            else:
                errors.append(f"包含'学员数据'的文件不唯一，且无唯一CSV：{[p.name for p in student_data_files]}")

    if errors:
        raise ValueError("\n".join(errors))

    assert chosen_student_file is not None

    return {

        "email": email_files[0],
        "mapping": mapping_files[0],
        "student": chosen_student_file,
    }


def read_csv_with_fallback(path: Path) -> pd.DataFrame:
    last_err = None
    for enc in ["utf-8-sig", "gb18030", "utf-8"]:
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
        except Exception as e:
            last_err = e
    raise ValueError(f"CSV读取失败: {path.name}, 错误: {last_err}")


def load_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return read_csv_with_fallback(path)
    return pd.read_excel(path, dtype=str)


def parse_email_table(path: Path) -> tuple[dict[str, str], list[str]]:
    errors = []

    # 邮箱格式验证正则
    email_pattern = re.compile(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')

    if path.suffix.lower() == ".csv":
        coach_to_email: dict[str, str] = {}
        coach_emails_set: dict[str, set[str]] = {}
        text = ""
        for enc in ["utf-8-sig", "gb18030", "utf-8"]:
            try:
                text = path.read_text(encoding=enc)
                break
            except Exception:
                continue
        if not text:
            raise ValueError(f"无法读取文件: {path}")

        for line in text.splitlines():
            line = clean_text(line)
            if not line or "教练" in line and "邮箱" in line:
                continue

            m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", line)
            email = m.group(0) if m else ""

            coach_part = line
            if email:
                coach_part = coach_part.replace(email, "")
            # 支持制表符和逗号两种分隔符格式
            coach_part = coach_part.replace("\t", " ").replace(",", " ")
            coach = clean_name(coach_part)
            if not coach:
                continue

            coach_emails_set.setdefault(coach, set())
            if email:
                coach_emails_set[coach].add(email)

        for coach, email_set in coach_emails_set.items():
            if len(email_set) > 1:
                errors.append(f"教练'{coach}'存在多个邮箱: {sorted(email_set)}")
            elif len(email_set) == 1:
                coach_to_email[coach] = next(iter(email_set))
            else:
                coach_to_email[coach] = ""
                errors.append(f"教练'{coach}'邮箱为空，跳过")

        return coach_to_email, errors

    # Excel 文件处理，添加异常处理
    try:
        df = load_table(path)
    except Exception as e:
        error_msg = str(e)
        if "BadZipFile" in error_msg or "not a zip file" in error_msg:
            raise ValueError(
                f"文件损坏或格式错误: {path.name}\n"
                f"  该Excel文件可能已损坏，请检查文件是否可以正常打开。\n"
                f"  建议：重新保存或导出该文件后再试。"
            )
        elif "No such file" in error_msg or "找不到" in error_msg:
            raise ValueError(f"文件不存在: {path}")
        else:
            raise ValueError(f"读取文件失败: {path.name}, 错误: {error_msg}")
    columns = [clean_text(c) for c in df.columns]
    coach_col = next((c for c in columns if "教练" in c), None)
    email_col = next((c for c in columns if "邮箱" in c), None)
    if not coach_col or not email_col:
        raise ValueError(f"教练邮箱表缺少关键列，当前列: {columns}")

    coach_to_email: dict[str, str] = {}
    coach_emails_set: dict[str, set[str]] = {}

    for _, row in df.iterrows():
        coach = clean_name(row.get(coach_col, ""))
        email = clean_text(row.get(email_col, ""))
        if not coach:
            continue
        coach_emails_set.setdefault(coach, set())
        if email:
            coach_emails_set[coach].add(email)

    for coach, email_set in coach_emails_set.items():
        if len(email_set) > 1:
            errors.append(f"教练'{coach}'存在多个邮箱: {sorted(email_set)}")
        elif len(email_set) == 1:
            email = next(iter(email_set))
            if not email_pattern.match(email):
                errors.append(f"教练'{coach}'邮箱格式异常: {email}，跳过")
            coach_to_email[coach] = email
        else:
            coach_to_email[coach] = ""
            errors.append(f"教练'{coach}'邮箱为空，跳过")

    return coach_to_email, errors


def get_col_by_keyword(df: pd.DataFrame, keyword: str) -> str | None:
    for c in df.columns:
        if keyword in clean_text(c):
            return c
    return None


def parse_student_names(cell_value: Any) -> list[str]:
    text = clean_text(cell_value)
    if not text:
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    raw_lines = text.split("\n")
    result = []
    for line in raw_lines:
        line = line.strip()
        if not line:
            continue
        # 连续2+空格视为多人分隔符（如"北巷往北走      李盛"）
        if "  " in line:
            parts = [clean_name(x) for x in re.split(r'\s{2,}', line)]
            result.extend([p for p in parts if p])
        else:
            cleaned = clean_name(line)
            if cleaned:
                result.append(cleaned)
    return result


def parse_mapping_table(path: Path) -> list[CoachAssignment]:
    """
    解析对照表，支持多sheet（如各期数据），从sheet名提取期数
    
    Sheet名格式示例：'第3期'、'3期'、'第5期数据' 等，提取其中的数字作为期数
    """
    all_assignments: list[CoachAssignment] = []
    
    try:
        if path.suffix.lower() in {".xlsx", ".xls"}:
            with pd.ExcelFile(path) as xl:
                sheet_names = xl.sheet_names  # 读取所有sheet
                
                for sheet_name in sheet_names:
                    # 从sheet名提取期数（如 '第3期' -> 3, '第5期数据' -> 5）
                    period_match = re.search(r'第?(\d+)期', sheet_name)
                    period_num = int(period_match.group(1)) if period_match else None
                    
                    df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str).fillna("")
                    
                    coach_col = get_col_by_keyword(df, "教练")
                    count_col = get_col_by_keyword(df, "安排学员数量")
                    students_col = get_col_by_keyword(df, "咨询学员名称")
                    ops_col = get_col_by_keyword(df, "运营")
                    qr_col = get_col_by_keyword(df, "对应运营二维码")

                    if not coach_col or not students_col:
                        continue  # 跳过缺少关键列的sheet

                    for _, row in df.iterrows():
                        coach = clean_name(row.get(coach_col, ""))
                        if not coach:
                            continue
                        planned_count = parse_int(row.get(count_col, "")) if count_col else None
                        students = parse_student_names(row.get(students_col, ""))

                        # 过滤运营"加分"干扰（如运营列"天天分1"→过滤学员名"天天分"）
                        if ops_col:
                            ops_text = clean_text(row.get(ops_col, ""))
                            ops_filter_set = set()
                            for part in re.split(r'[\n\r]+', ops_text):
                                part = part.strip()
                                if not part:
                                    continue
                                for sub in part.split():
                                    sub = sub.strip()
                                    if sub:
                                        ops_filter_set.add(sub)
                                        # 去掉"分1""分3"等后缀得到基础名
                                        base = re.sub(r'分\d+$', '', sub).strip()
                                        if base and base != sub:
                                            ops_filter_set.add(base)
                                            ops_filter_set.add(base + "分")
                            students = [s for s in students if s not in ops_filter_set]

                        # 自动补入"对应运营二维码"列中的学员名
                        if qr_col:
                            qr_names = parse_student_names(row.get(qr_col, ""))
                            students.extend(qr_names)

                        all_assignments.append(CoachAssignment(
                            coach=coach, 
                            planned_count=planned_count, 
                            student_names=students,
                            period=period_num
                        ))
        else:
            # CSV文件无sheet概念，期数为None
            df = load_table(path).fillna("")
            
            coach_col = get_col_by_keyword(df, "教练")
            count_col = get_col_by_keyword(df, "安排学员数量")
            students_col = get_col_by_keyword(df, "咨询学员名称")
            ops_col = get_col_by_keyword(df, "运营")
            qr_col = get_col_by_keyword(df, "对应运营二维码")

            if not coach_col or not students_col:
                raise ValueError(f"对照表缺少关键列，当前列: {[clean_text(c) for c in df.columns]}")

            for _, row in df.iterrows():
                coach = clean_name(row.get(coach_col, ""))
                if not coach:
                    continue
                planned_count = parse_int(row.get(count_col, "")) if count_col else None
                students = parse_student_names(row.get(students_col, ""))

                if ops_col:
                    ops_text = clean_text(row.get(ops_col, ""))
                    ops_filter_set = set()
                    for part in re.split(r'[\n\r]+', ops_text):
                        part = part.strip()
                        if not part:
                            continue
                        for sub in part.split():
                            sub = sub.strip()
                            if sub:
                                ops_filter_set.add(sub)
                                base = re.sub(r'分\d+$', '', sub).strip()
                                if base and base != sub:
                                    ops_filter_set.add(base)
                                    ops_filter_set.add(base + "分")
                    students = [s for s in students if s not in ops_filter_set]

                if qr_col:
                    qr_names = parse_student_names(row.get(qr_col, ""))
                    students.extend(qr_names)

                all_assignments.append(CoachAssignment(
                    coach=coach, 
                    planned_count=planned_count, 
                    student_names=students,
                    period=None
                ))
                
    except Exception as e:
        error_msg = str(e)
        if "BadZipFile" in error_msg or "not a zip file" in error_msg:
            raise ValueError(
                f"对照表文件损坏或格式错误: {path.name}\n"
                f"  该Excel文件可能已损坏，请检查文件是否可以正常打开。\n"
                f"  建议：重新保存或导出该文件后再试。"
            )
        else:
            raise ValueError(f"读取对照表失败: {path.name}, 错误: {error_msg}")

    return all_assignments


def dedup_student_table(path: Path) -> dict[str, StudentInfo]:
    try:
        df = load_table(path).fillna("")
    except Exception as e:
        error_msg = str(e)
        if "BadZipFile" in error_msg or "not a zip file" in error_msg:
            raise ValueError(
                f"学员数据文件损坏或格式错误: {path.name}\n"
                f"  该Excel文件可能已损坏，请检查文件是否可以正常打开。\n"
                f"  建议：重新保存或导出该文件后再试。"
            )
        else:
            raise ValueError(f"读取学员数据表失败: {path.name}, 错误: {error_msg}")

    nickname_col = get_col_by_keyword(df, "微信昵称")
    if not nickname_col:
        raise ValueError(f"学员数据表缺少'微信昵称'列，当前列: {[clean_text(c) for c in df.columns]}")

    date_col = get_col_by_keyword(df, "购买日期")

    needed_keywords = [
        "手机号", "期数", "工作状态", "待业时长", "工作年限", "行业岗位", "当前薪资区间", "工作阶段", "城市",
        "心力状态", "理想职业", "过去半年自我能力提升投入", "过往自我投入领域", "过往投入其他说明", "对优势星球了解程度", "问题描述"
    ]

    col_map = {k: get_col_by_keyword(df, k) for k in needed_keywords}

    if date_col:
        df["__parsed_date"] = df[date_col].apply(parse_date)
        df = df.sort_values(by=["__parsed_date"], ascending=True, na_position="first")

    result: dict[str, StudentInfo] = {}
    for _, row in df.iterrows():
        nickname = clean_name(row.get(nickname_col, ""))
        if not nickname:
            continue
        info = {
            "微信昵称": nickname,
            "手机号": clean_text(row.get(col_map["手机号"], "")) if col_map["手机号"] else "",
            "期数": clean_text(row.get(col_map["期数"], "")) if col_map["期数"] else "",
            "工作状态": clean_text(row.get(col_map["工作状态"], "")) if col_map["工作状态"] else "",
            "待业时长": clean_text(row.get(col_map["待业时长"], "")) if col_map["待业时长"] else "",
            "工作年限": clean_text(row.get(col_map["工作年限"], "")) if col_map["工作年限"] else "",
            "行业岗位": clean_text(row.get(col_map["行业岗位"], "")) if col_map["行业岗位"] else "",
            "当前薪资区间": clean_text(row.get(col_map["当前薪资区间"], "")) if col_map["当前薪资区间"] else "",
            "工作阶段": clean_text(row.get(col_map["工作阶段"], "")) if col_map["工作阶段"] else "",
            "城市": clean_text(row.get(col_map["城市"], "")) if col_map["城市"] else "",
            "心力状态": clean_text(row.get(col_map["心力状态"], "")) if col_map["心力状态"] else "",
            "理想职业": clean_text(row.get(col_map["理想职业"], "")) if col_map["理想职业"] else "",
            "过去半年自我能力提升投入": clean_text(row.get(col_map["过去半年自我能力提升投入"], "")) if col_map["过去半年自我能力提升投入"] else "",
            "过往自我投入领域": clean_text(row.get(col_map["过往自我投入领域"], "")) if col_map["过往自我投入领域"] else "",
            "过往投入其他说明": clean_text(row.get(col_map["过往投入其他说明"], "")) if col_map["过往投入其他说明"] else "",
            "对优势星球了解程度": clean_text(row.get(col_map["对优势星球了解程度"], "")) if col_map["对优势星球了解程度"] else "",
            "问题描述": clean_text(row.get(col_map["问题描述"], "")) if col_map["问题描述"] else "",
        }
        result[nickname] = info
    return result


def edit_distance(s1: str, s2: str) -> int:
    """计算两个字符串的Levenshtein编辑距离"""
    if len(s1) < len(s2):
        return edit_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = prev_row[j + 1] + 1
            deletions = curr_row[j] + 1
            substitutions = prev_row[j] + (c1 != c2)
            curr_row.append(min(insertions, deletions, substitutions))
        prev_row = curr_row
    return prev_row[-1]


def fuzzy_match_student(query_name: str, student_map: dict[str, StudentInfo], search_names: set[str]) -> tuple[str | None, list[str]]:
    """
    三级模糊匹配（在指定的学员范围内查找）

    参数:
        query_name: 待查询的学员名
        student_map: 全部学员信息映射
        search_names: 模糊匹配的搜索范围（已按期数筛选）

    返回:
        - 匹配到的微信昵称（唯一匹配时），或 None
        - 候选列表：唯一匹配时为空，不唯一时为候选名列表，无匹配时为空
    """
    # 级别2：部分包含匹配
    partial_matches = []
    for nickname in search_names:
        if query_name in nickname or nickname in query_name:
            common_len = min(len(query_name), len(nickname))
            if common_len >= 3:
                partial_matches.append(nickname)
    if len(partial_matches) == 1:
        return partial_matches[0], []
    if len(partial_matches) > 1:
        return None, partial_matches

    # 级别3：编辑距离匹配 —— 字符数相同，编辑距离 ≤1
    ed_matches = []
    for nickname in search_names:
        if len(query_name) == len(nickname):
            dist = edit_distance(query_name, nickname)
            if dist <= 1:
                ed_matches.append(nickname)
    if len(ed_matches) == 1:
        return ed_matches[0], []
    if len(ed_matches) > 1:
        return None, ed_matches

    return None, []


def build_results(assignments: list[CoachAssignment], coach_email_map: dict[str, str], student_map: dict[str, StudentInfo]) -> tuple[list[CoachResult], list[str], list[str]]:
    warnings: list[str] = []
    errors: list[str] = []
    results: list[CoachResult] = []

    # 预处理：按期数分组学员昵称（用于模糊匹配范围限定）
    # period_students_map: 期数 -> 该期学员昵称集合
    period_students_map: dict[int | None, set[str]] = {}
    for nickname, info in student_map.items():
        period_num = parse_int(info.get("期数", ""))
        if period_num not in period_students_map:
            period_students_map[period_num] = set()
        period_students_map[period_num].add(nickname)
    
    # 全部学员集合（作为fallback）
    all_student_names = set(student_map.keys())

    for item in assignments:
        # 只处理在教练邮箱表中存在的教练
        if item.coach not in coach_email_map:
            continue

        email = coach_email_map.get(item.coach, "")
        parsed_count = len(item.student_names)

        # 根据assignment的period确定模糊匹配范围
        # 如果有明确期数，在该期学员中查找；否则在全部学员中查找
        if item.period is not None and item.period in period_students_map:
            fuzzy_search_names = period_students_map[item.period]
        else:
            fuzzy_search_names = all_student_names

        matched = []
        missing = []
        fuzzy_nonunique = []

        for name in item.student_names:
            if name in student_map:
                matched.append(student_map[name])
            else:
                # 尝试模糊匹配（在对应期数或全部学员中）
                fuzzy_nick, match_candidates = fuzzy_match_student(name, student_map, fuzzy_search_names)
                if fuzzy_nick and fuzzy_nick in student_map:
                    matched.append(student_map[fuzzy_nick])
                elif match_candidates:
                    fuzzy_nonunique.append(f"'{name}' → {match_candidates}")
                    missing.append(name)
                else:
                    missing.append(name)

        sendable = True
        if not email:
            sendable = False
            errors.append(f"【{item.coach}】教练 未配置邮箱，跳过发送")

        if missing:
            # 每个学员名单独用【】包裹
            missing_bracketed = "、".join(f"【{name}】" for name in missing)
            missing_names = "、".join(missing)
            warnings.append(f"分配给【{item.coach}】教练的学员{missing_bracketed}，在学员数据表中未找到（已跳过，请手动核实处理）：{missing_names}")

        if fuzzy_nonunique:
            # 格式化模糊匹配提示
            fuzzy_details = "、".join(fuzzy_nonunique)
            warnings.append(f"【{item.coach}】教练 模糊匹配不唯一：{fuzzy_details}")

        results.append(
            CoachResult(
                coach=item.coach,
                email=email,
                planned_count=item.planned_count,
                parsed_count=parsed_count,
                matched_count=len(matched),
                missing_students=missing,
                students=matched,
                sendable=sendable,
            )
        )

    return results, warnings, errors


def export_check_excel(date_str: str, results: list[CoachResult], output_dir: Path) -> Path:
    """
    将待发送的教练+学员信息导出为Excel，方便检查和留底。
    本质上是从学员数据表中截取匹配行，再加一列教练名称。
    """
    # 定义列顺序：教练名 + 学员数据表原始字段
    student_fields = [
        "微信昵称", "手机号", "期数", "工作状态", "待业时长", "工作年限",
        "行业岗位", "当前薪资区间", "工作阶段", "城市", "心力状态",
        "理想职业", "过去半年自我能力提升投入", "过往自我投入领域",
        "过往投入其他说明", "对优势星球了解程度", "问题描述",
    ]
    columns = ["教练"] + student_fields

    rows = []
    for r in results:
        for stu in r.students:
            row = {"教练": r.coach}
            for field in student_fields:
                row[field] = stu.get(field, "")
            rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    output_path = output_dir / f"发送检查_{date_str}.xlsx"
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


def build_preview_text(date_str: str, result: CoachResult) -> str:
    lines = [
        f"日期：{date_str}",
        f"教练：{result.coach}",
        f"邮箱：{result.email or '（空）'}",
        f"对接的学员人数：{result.matched_count}",
        "---",
    ]

    for i, stu in enumerate(result.students, start=1):
        lines.extend(
            [
                f"({i})",
                f"学员微信昵称：{stu.get('微信昵称', '')}",
                f"手机号：{stu.get('手机号', '')}",
                f"期数：{stu.get('期数', '')}",
                "",
            ]
        )
    return "\n".join(lines).strip() + "\n"


def get_display_value(value: str) -> str:
    """获取显示值，空值显示为（未填写）"""
    return value if value and value.strip() else "（未填写）"


def build_email_body(date_str: str, result: CoachResult) -> str:
    lines = [
        f"#{date_str}",
        f"{result.coach}教练你好！你本周对接的学员数量为 {result.matched_count}，详细情况如下：",
        "",
    ]

    for i, stu in enumerate(result.students, start=1):
        lines.extend(
            [
                f"（{i}）",
                f"- 学员微信昵称：{get_display_value(stu.get('微信昵称', ''))}",
                f"- 手机号：{get_display_value(stu.get('手机号', ''))}",
                f"- 期数：{get_display_value(stu.get('期数', ''))}",
                "",
                f"- 工作状态：{get_display_value(stu.get('工作状态', ''))}",
                f"- 待业时长：{get_display_value(stu.get('待业时长', ''))}",
                f"- 工作年限：{get_display_value(stu.get('工作年限', ''))}",
                f"- 行业岗位：{get_display_value(stu.get('行业岗位', ''))}",
                f"- 当前薪资区间：{get_display_value(stu.get('当前薪资区间', ''))}",
                f"- 工作阶段：{get_display_value(stu.get('工作阶段', ''))}",
                f"- 城市：{get_display_value(stu.get('城市', ''))}",
                f"- 心力状态：{get_display_value(stu.get('心力状态', ''))}",
                f"- 理想职业：{get_display_value(stu.get('理想职业', ''))}",
                f"- 过去半年自我能力提升投入：{get_display_value(stu.get('过去半年自我能力提升投入', ''))}",
                f"- 过往自我投入领域：{get_display_value(stu.get('过往自我投入领域', ''))}",
                f"- 过往投入其他说明：{get_display_value(stu.get('过往投入其他说明', ''))}",
                f"- 对优势星球了解程度：{get_display_value(stu.get('对优势星球了解程度', ''))}",
                f"- 问题描述：{get_display_value(stu.get('问题描述', ''))}",
                "",
                "======================",
                "",
            ]
        )

    return "\n".join(lines).strip() + "\n"


def send_emails(
    sender_email: str,
    auth_code: str,
    smtp_host: str,
    smtp_port: int,
    date_str: str,
    results: list[CoachResult],
    sender_name: str,
) -> tuple[list[str], list[str]]:
    """发送邮件，支持智能重试和自动重连机制"""
    success_logs = []
    fail_logs = []
    
    # 错误记录：教练 -> 最后一次错误信息
    error_map: dict[str, str] = {}
    
    # 单次连接最多发送数量（163邮箱限制约10封，保守设为8）
    MAILS_PER_CONNECTION = 8
    RECONNECT_DELAY = 3  # 重连等待秒数

    sendable_results = [r for r in results if r.sendable]
    if not sendable_results:
        return success_logs, ["没有可发送的教练邮件"]

    # 预构建所有邮件
    email_cache: dict[str, tuple] = {}  # coach -> (msg, email)
    for r in sendable_results:
        subject = f"[{date_str}] L2学员对接明细 - {r.coach}"
        body = build_email_body(date_str, r)
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = str(Header(subject, "utf-8"))
        msg["From"] = formataddr((str(Header(sender_name, "utf-8")), sender_email))
        msg["To"] = str(r.email)
        email_cache[r.coach] = (msg, r.email)

    # 待发送队列
    pending = list(sendable_results)
    max_rounds = 2  # 最多重试1次，用于处理偶发网络抖动
    retry_delay = 10  # 重试等待时间

    for round_num in range(1, max_rounds + 1):
        if not pending:
            break  # 全部发送成功

        failed_this_round = []
        sent_count = 0  # 当前连接已发送数量
        
        # 建立新连接
        server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
        try:
            server.login(sender_email, auth_code)
            
            for i, r in enumerate(pending):
                msg, email = email_cache[r.coach]
                
                # 检查是否需要重连
                if sent_count >= MAILS_PER_CONNECTION:
                    server.quit()
                    time.sleep(RECONNECT_DELAY)
                    server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
                    server.login(sender_email, auth_code)
                    sent_count = 0
                
                try:
                    server.sendmail(sender_email, [email], msg.as_string())
                    success_logs.append(f"发送成功: 教练[{r.coach}] -> {email}")
                    print(f"[OK] {r.coach} -> {email}")
                    sent_count += 1
                    time.sleep(0.2)  # 成功后短暂等待
                    
                except Exception as e:
                    failed_this_round.append(r)
                    error_map[r.coach] = str(e)
                    print(f"[失败] {r.coach}: {e}")
        finally:
            try:
                server.quit()
            except:
                pass

        # 更新待发送队列
        pending = failed_this_round

        # 如果还有失败的，尝试重试一次
        if pending and round_num < max_rounds:
            print(f"\n{len(pending)} 封失败，{retry_delay}秒后重试...")
            time.sleep(retry_delay)
    
    # 最终还失败的，记录日志
    if pending:
        for r in pending:
            fail_logs.append(f"发送失败: 教练[{r.coach}] -> {r.email}, 错误: {error_map[r.coach]}")

    return success_logs, fail_logs


def write_log(path: Path, lines: list[str]) -> None:
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def delete_table_file(path: Path) -> bool:
    """
    删除表格文件
    
    参数:
        path: 表格文件路径
    
    返回:
        True 表示成功，False 表示失败
    """
    try:
        if path.exists():
            path.unlink()
            return True
        else:
            return False
    except Exception as e:
        print(f"删除文件失败: {path.name}, 错误: {e}")
        return False


def write_csv_log(
    root_dir: Path,
    timestamp: str,
    results: list[CoachResult],
    send_success: dict[str, bool],
    fail_reasons: dict[str, str],  # 新增：教练 -> 失败原因
    warnings: list[str],
    errors: list[str],
) -> Path:
    """写入统一CSV日志文件，返回日志文件路径"""
    log_path = root_dir / "log.csv"
    
    # 读取现有内容获取当前序号
    existing_rows: list[list[str]] = []
    if log_path.exists():
        with open(log_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # 跳过表头
            existing_rows = [row for row in reader if row]
    
    start_index = len(existing_rows) + 1
    
    # 构建错误信息映射（教练 -> 异常信息列表）
    coach_errors: dict[str, list[str]] = {}
    for e in errors:
        for r in results:
            if r.coach in e:
                coach_errors.setdefault(r.coach, []).append(e)
                break
    
    # 构建警告信息映射（教练 -> 警告信息列表）
    coach_warnings: dict[str, list[str]] = {}
    for w in warnings:
        for r in results:
            if r.coach in w:
                coach_warnings.setdefault(r.coach, []).append(w)
                break
    
    # 准备新行数据
    new_rows: list[list[str]] = []
    for i, r in enumerate(results, start=start_index):
        # 咨询学员名称（回车分割）
        student_names = "\n".join([s.get("微信昵称", "") for s in r.students])
        
        # 发送情况
        if r.coach in send_success:
            send_status = "成功" if send_success[r.coach] else "失败"
        elif not r.sendable:
            send_status = "跳过（无邮箱）"
        else:
            send_status = "未发送"
        
        # 异常情况
        exceptions: list[str] = []
        # 发送失败原因
        if r.coach in fail_reasons:
            exceptions.append(f"发送失败: {fail_reasons[r.coach]}")
        if r.coach in coach_errors:
            exceptions.extend(coach_errors[r.coach])
        if r.coach in coach_warnings:
            exceptions.extend(coach_warnings[r.coach])
        if r.missing_students:
            exceptions.append(f"未找到学员: {r.missing_students}")
        exception_text = "\n".join(exceptions) if exceptions else ""
        
        row = [
            str(i),  # 序号
            timestamp,  # 时间
            r.coach,  # 教练
            str(r.matched_count),  # 对接的学员人数
            student_names,  # 咨询学员名称（回车分割）
            r.email or "",  # 邮箱
            send_status,  # 发送情况
            exception_text,  # 异常情况
        ]
        new_rows.append(row)
    
    # 写入文件
    file_exists = log_path.exists()
    with open(log_path, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["序号", "时间", "教练", "对接的学员人数", "咨询学员名称", "邮箱", "发送情况", "异常情况"])
        writer.writerows(new_rows)
        # 添加分割行，标识本次批次结束
        writer.writerow(["---", "---", "---", "---", "---", "---", "---", "---"])
    
    return log_path


def read_sender_config(config_path: Path) -> SenderConfig | None:
    """从配置文件读取发件人信息，文件不存在返回None
    
    如果检测到授权码为明文，会自动编码并更新配置文件
    """
    if not config_path.exists():
        return None
    
    content = config_path.read_text(encoding="utf-8")
    config: dict[str, Any] = {
        "sender_email": "",
        "auth_code": "",
        "smtp_host": "",
        "smtp_port": 465,
        "sender_name": "",
        "data_dir": ""
    }
    
    auth_code_plain = None  # 用于记录明文授权码
    
    for line in content.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if key == "smtp_port":
                try:
                    config[key] = int(value)
                except ValueError:
                    pass
            elif key == "auth_code":
                # 检查是否为明文
                if is_auth_code_plain(value):
                    auth_code_plain = value  # 记录明文
                    config[key] = value  # 暂时保存明文，后面会编码
                else:
                    # 已编码，解码后使用
                    config[key] = decode_auth_code(value)
            else:
                config[key] = value
    
    # 如果发现明文授权码，编码后更新配置文件
    if auth_code_plain is not None:
        encoded = encode_auth_code(auth_code_plain)
        config["auth_code"] = auth_code_plain  # 返回明文供本次使用
        # 更新配置文件
        _update_auth_code_in_config(config_path, encoded)
        print(f"[安全提示] 检测到明文授权码，已自动编码并更新配置文件")
    
    return SenderConfig(**config)


def _update_auth_code_in_config(config_path: Path, encoded_auth: str) -> None:
    """更新配置文件中的授权码为编码后的值"""
    content = config_path.read_text(encoding="utf-8")
    lines = content.splitlines()
    new_lines = []
    
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("auth_code") and "=" in stripped:
            # 替换授权码行
            key = line.split("=")[0]
            indent = line[:len(line) - len(line.lstrip())]  # 保留原始缩进
            new_lines.append(f"{indent}auth_code = {encoded_auth}")
        else:
            new_lines.append(line)
    
    config_path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")


def save_sender_config(config_path: Path, config: SenderConfig) -> None:
    """保存发件人配置到文件（授权码会自动编码）"""
    # 编码授权码
    encoded_auth = encode_auth_code(config.auth_code) if config.auth_code else ""
    
    content = f"""# 发件人配置
# 发送邮箱账号
sender_email = {config.sender_email}
# SMTP授权码（已编码，请勿修改）
auth_code = {encoded_auth}
# SMTP服务器地址
smtp_host = {config.smtp_host}
# SMTP端口（SSL端口一般为465）
smtp_port = {config.smtp_port}
# 发件人显示名称（邮件Header中的From名称）
sender_name = {config.sender_name}
# 数据目录（包含三个表格文件和sender_config.txt的目录）
data_dir = {config.data_dir}
"""
    config_path.write_text(content, encoding="utf-8")


def is_double_click_run() -> bool:
    """检测是否是通过双击运行的"""
    # Windows双击运行时，stdin通常不是终端
    # 或者检查是否在IDE中运行
    return not sys.stdin.isatty() or (os.environ.get("TERM_PROGRAM") is None and os.environ.get("VSCODE_PID") is None)


def show_program_info():
    """显示程序功能说明"""
    print("=" * 70)
    print("                    优势星球 L2 学员信息自动分发工具")
    print("=" * 70)
    print()
    print("【程序功能】")
    print("  本程序用于自动将学员信息分发到对应的教练邮箱。快速，高效！")
    print()
    print("=" * 70)


def prompt_period_confirmation(available_periods: list[int]) -> int | None:
    """
    询问用户确认期数或输入新期数
    
    参数:
        available_periods: 对照表中可用的期数列表
    
    返回:
        用户确认的期数，或 None 表示取消
    """
    if TEST_MODE:
        print("\n[测试模式] 自动确认期数")
        return available_periods[0] if available_periods else None
    
    # 默认取最新一期（最大期数）
    default_period = max(available_periods) if available_periods else None
    available_periods_str = "、".join(str(p) for p in sorted(available_periods))
    
    while True:
        # 分两行输出，美观
        print(f"\n即将处理第【{default_period}】期的数据。如果期数没问题，请输入 Y ；")
        answer = input("如果要指定期数，请输入你想处理的期数（只输入数字）： ").strip()
        
        # 输入 y 确认
        if answer.lower() in ("y", "yes", "是"):
            return default_period
        
        # 输入 n 取消
        if answer.lower() in ("n", "no", "否"):
            return None
        
        # 尝试解析为数字
        try:
            period_input = int(answer)
            if period_input in available_periods:
                return period_input
            else:
                print(f"  期数【{period_input}】不在对照表中，可用期数：{available_periods_str}")
        except ValueError:
            print("  请输入 Y 确认，或输入期数数字（如：84）")


def prompt_send_confirmation(has_warnings: bool = False) -> bool:
    """询问用户是否开始发送"""
    if TEST_MODE:
        print("\n[测试模式] 自动确认开始发送邮件")
        return True
    while True:
        if has_warnings:
            answer = input("\n确认发送邮件请输入 Y ，取消请输入 N（建议处理异常后再发送邮件）：").strip().lower()
        else:
            answer = input("\n确认发送邮件请输入 Y ，取消请输入 N ：").strip().lower()
        if answer in ("y", "yes", "是"):
            return True
        if answer in ("n", "no", "否"):
            return False
        print("请输入 Y 或 N")


def prompt_for_config() -> SenderConfig:
    """引导用户交互式输入配置信息"""
    print()
    print("【首次使用配置】")
    print("请按提示输入发件人配置信息：")
    print()
    
    # 必填项
    while True:
        sender_email = input("请输入发送邮箱: ").strip()
        if sender_email:
            break
        print("发送邮箱不能为空，请重新输入")
    
    auth_code = getpass.getpass("请输入SMTP授权码(输入过程不可见): ").strip()
    
    while True:
        smtp_host = input("请输入SMTP服务器地址: ").strip()
        if smtp_host:
            break
        print("SMTP服务器地址不能为空，请重新输入")
    
    # 端口（有默认值）
    smtp_port = 465
    try:
        port_input = input("请输入SMTP端口(默认: 465): ").strip()
        if port_input:
            smtp_port = int(port_input)
    except ValueError:
        print("端口输入无效，使用默认值 465")
    
    # 可选项（需用户输入）
    while True:
        sender_name = input("请输入发件人显示名称: ").strip()
        if sender_name:
            break
        print("发件人显示名称不能为空，请重新输入")
    
    while True:
        data_dir = input("请输入数据目录: ").strip()
        if data_dir:
            break
        print("数据目录不能为空，请重新输入")
    
    return SenderConfig(
        sender_email=sender_email,
        auth_code=auth_code,
        smtp_host=smtp_host,
        smtp_port=smtp_port,
        sender_name=sender_name,
        data_dir=data_dir
    )


def ensure_config_exists(config_path: Path) -> SenderConfig:
    """确保配置文件存在，不存在则引导用户创建"""
    if config_path.exists():
        config = read_sender_config(config_path)
        if config is not None:
            return config
    
    print(f"\n未找到配置文件: {config_path}")
    config = prompt_for_config()
    
    save_sender_config(config_path, config)
    print(f"\n配置已保存到: {config_path}")
    
    return config


def run_interactive(args):
    """交互式运行流程"""
    # 1. 显示程序信息
    show_program_info()
    
    # 2. 从程序根目录读取配置
    config_path = get_app_dir() / "sender_config.txt"
    config = ensure_config_exists(config_path)
    
    # 3. 获取数据目录
    data_dir = Path(config.data_dir).resolve()
    # print(f"\n使用数据目录: {data_dir}")
    
    if not data_dir.exists() or not data_dir.is_dir():
        print(f"\n错误: 数据目录不存在: {data_dir}")
        input("按回车键退出...")
        raise ValueError(f"数据目录不存在: {data_dir}")
    
    if not config.auth_code:
        print("\n警告: 授权码未配置")
        config.auth_code = getpass.getpass("请输入SMTP授权码(输入过程不可见): ").strip()
        if config.auth_code:
            # 询问是否保存到根目录配置
            save_auth = input("是否保存授权码到配置文件？(y/n): ").strip().lower()
            if save_auth in ("y", "yes", "是"):
                save_sender_config(config_path, config)
                print("授权码已保存")
    
    # 4. 检测文件（成功不提示，失败才提示）
    try:
        files = detect_files(data_dir)
    except ValueError as e:
        print("\n文件识别失败：")
        print(e)
        input("\n按回车键退出...")
        raise
    
    # 5. 解析对照表获取期数信息
    assignments = parse_mapping_table(files["mapping"])
    
    # 提取所有期数
    available_periods = sorted(set(a.period for a in assignments if a.period is not None))
    
    # 6. 询问用户确认期数
    confirmed_period = prompt_period_confirmation(available_periods)
    if confirmed_period is None:
        print("\n已取消执行。程序退出。")
        return
    
    # 根据确认的期数筛选assignments
    if confirmed_period in available_periods:
        assignments = [a for a in assignments if a.period == confirmed_period]
    
    # 7. 解析其他数据
    coach_email_map, email_table_errors = parse_email_table(files["email"])
    student_map = dedup_student_table(files["student"])
    
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

    results, warnings, errors = build_results(assignments, coach_email_map, student_map)
    # 不再添加邮箱错误到errors（取消邮箱错误提醒）
    # errors.extend(email_table_errors)

    # 8. 生成检查用Excel
    root_dir = get_app_dir().resolve()
    check_excel_path = export_check_excel(date_str, results, root_dir)
    # print(f"\n检查用Excel已生成: {check_excel_path}")

    # 9. 显示预览（内存中生成，不写入文件）
    # print("\n" + "=" * 70)
    # print("正在生成预览...")
    
    preview_parts = []
    for r in results:
        preview_parts.append(build_preview_text(date_str, r))
    
    # 10. 显示预览摘要（仅显示有邮箱的条目）
    print("\n" + "=" * 70)
    print("发送预览摘要：")
    sendable_results = [r for r in results if r.email and str(r.email).strip() and str(r.email).lower() != 'nan']
    for i, r in enumerate(sendable_results, 1):
        print(f"  {i}. {r.coach}: {r.matched_count}人 -> {r.email}")
    
    # 11. 显示数据异常提醒（如有）
    if warnings:
        print("\n" + "=" * 70)
        print('\n【数据异常提醒】未找到以下学员，请根据下面的提示核实优化"学员对照表"：\n')
        for w in warnings:
            print(f"  ! {w}")
    
    if args.preview_only:
        print("\n[仅预览模式] 已生成预览，未发送邮件。")
        input("\n按回车键退出...")
        return
    
    # 12. 询问是否发送
    print("\n" + "=" * 70)
    has_warnings = bool(warnings)
    if not prompt_send_confirmation(has_warnings):
        print("\n已取消发送。")
        input("\n按回车键退出...")
        return
    
    # 13. 执行发送（只发送有邮箱的）
    print("\n" + "=" * 70)
    print("开始发送邮件...")
    
    if not config.auth_code:
        print("错误: 未提供SMTP授权码，无法发送")
        input("\n按回车键退出...")
        raise ValueError("未提供SMTP授权码")
    
    # 过滤出有邮箱的教练
    sendable_results = [r for r in results if r.email and str(r.email).strip() and str(r.email).lower() != 'nan']
    
    if not sendable_results:
        print("没有需要发送邮件的教练（所有教练均无邮箱）。")
        input("\n按回车键退出...")
        return
    
    success, fail = send_emails(
        sender_email=config.sender_email,
        auth_code=config.auth_code,
        smtp_host=config.smtp_host,
        smtp_port=config.smtp_port,
        date_str=date_str,
        results=sendable_results,  # 只传入有邮箱的
        sender_name=config.sender_name,
    )
    
    # 14. 构建发送结果映射并保存CSV日志
    send_success: dict[str, bool] = {}
    fail_reasons: dict[str, str] = {}
    
    for s in success:
        coach = s.replace("发送成功: 教练[", "").split("]")[0] if "教练[" in s else ""
        if coach:
            send_success[coach] = True
    for f in fail:
        coach = f.replace("发送失败: 教练[", "").split("]")[0] if "教练[" in f else ""
        if coach:
            send_success[coach] = False
            if ", 错误: " in f:
                reason = f.split(", 错误: ")[-1]
                fail_reasons[coach] = reason
    
    # 获取根目录（脚本所在目录）
    root_dir = get_app_dir().resolve()
    log_path = write_csv_log(root_dir, timestamp, results, send_success, fail_reasons, warnings, errors)
    
    # 15. 显示发送结果
    print("\n" + "=" * 70)
    print("发送完成！")
    print()
    print(f"  成功: {len(success)} 封")
    if fail:
        print(f"  失败: {len(fail)} 封")
    print()
    print(f"  CSV日志: {log_path}")
    
    # 16. 清除表格数据
    print("\n" + "=" * 70)
    clear_input = input("按回车键删除【学员数据表】和【学员对接表】（输入其他任意键跳过）...")
    if clear_input.strip() == "":
        student_deleted = delete_table_file(files["student"])
        mapping_deleted = delete_table_file(files["mapping"])
        
        print()
        if student_deleted and mapping_deleted:
            print("  表格文件已删除。")
        else:
            if not student_deleted:
                print(f"  学员数据表删除失败: {files['student'].name}")
            if not mapping_deleted:
                print(f"  学员对接表删除失败: {files['mapping'].name}")
    
    print("\n" + "=" * 70)
    input("按回车键退出...")


def run(args):
    """主运行函数"""
    # 检测是否双击运行，如果是则使用交互式流程
    if is_double_click_run() or args.interactive:
        run_interactive(args)
    else:
        # 命令行模式（保持原有行为）
        run_cli(args)


def run_cli(args):
    """命令行模式运行（保持原有行为）"""
    # 从程序根目录读取配置文件
    config_path = get_app_dir() / "sender_config.txt"
    config: SenderConfig | None = None
    
    if config_path.exists():
        config = read_sender_config(config_path)
    
    # 如果仍然没有配置，使用命令行参数或提示输入
    if config is None:
        # 检查命令行参数是否完整
        if not args.sender_email or not args.smtp_host:
            print("错误: 未找到配置文件，且命令行参数不完整")
            print("请提供以下参数：--sender-email, --smtp-host")
            print("或创建 sender_config.txt 配置文件")
            print(f"配置文件路径: {config_path}")
            sys.exit(1)
        
        auth_code = args.auth_code or os.getenv("SMTP_AUTH_CODE", "")
        if not auth_code:
            auth_code = getpass.getpass("请输入SMTP授权码(输入过程不可见): ").strip()
        
        config = SenderConfig(
            sender_email=args.sender_email,
            auth_code=auth_code,
            smtp_host=args.smtp_host,
            smtp_port=args.smtp_port,
            sender_name=args.sender_name,
            data_dir=args.data_dir
        )
    
    # 获取数据目录
    if args.data_dir:
        data_dir = Path(args.data_dir).resolve()
    elif config.data_dir:
        data_dir = Path(config.data_dir).resolve()
    else:
        raise ValueError("未指定数据目录，请通过 --data-dir 参数或配置文件设置")

    if not data_dir.exists() or not data_dir.is_dir():
        raise ValueError(f"数据目录不存在: {data_dir}")


    files = detect_files(data_dir)
    coach_email_map, email_table_errors = parse_email_table(files["email"])
    assignments = parse_mapping_table(files["mapping"])
    student_map = dedup_student_table(files["student"])

    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

    results, warnings, errors = build_results(assignments, coach_email_map, student_map)
    errors.extend(email_table_errors)

    root_dir = get_app_dir().resolve()
    check_excel_path = export_check_excel(date_str, results, root_dir)
    print(f"检查用Excel已生成: {check_excel_path}")

    print("=" * 70)
    print("文件识别结果：")
    print(f"- 表1(邮箱): {files['email'].name}")
    print(f"- 表2(对照表): {files['mapping'].name}")
    print(f"- 表3(学员数据): {files['student'].name}")
    print("=" * 70)
    print("可发送教练数量：", len([r for r in results if r.sendable]))
    print("跳过教练数量：", len([r for r in results if not r.sendable]))

    if warnings:
        print("\n[一致性提醒]")
        for w in warnings:
            print("-", w)

    if errors:
        print("\n[错误提醒]")
        for e in errors:
            print("-", e)

    if args.preview_only:
        print("\n已启用 --preview-only，仅预览，不发送邮件。")
        return

    if TEST_MODE:
        print("\n[测试模式] 自动确认发送邮件")
    else:
        answer = input("\n是否确认发送邮件？输入 Y 确认，其它任意键取消: ").strip().lower()
        if answer != "y":
            print("已取消发送。")
            return

    # 如果没有授权码，尝试从环境变量或提示输入
    if not config.auth_code:
        config.auth_code = getpass.getpass("请输入SMTP授权码(输入过程不可见): ").strip()

    if not config.auth_code:
        raise ValueError("未提供SMTP授权码，无法发送")

    success, fail = send_emails(
        sender_email=config.sender_email,
        auth_code=config.auth_code,
        smtp_host=config.smtp_host,
        smtp_port=config.smtp_port,
        date_str=date_str,
        results=results,
        sender_name=config.sender_name,
    )

    # 构建发送结果映射并保存CSV日志
    send_success: dict[str, bool] = {}
    fail_reasons: dict[str, str] = {}
    
    for s in success:
        coach = s.replace("发送成功: 教练[", "").split("]")[0] if "教练[" in s else ""
        if coach:
            send_success[coach] = True
    for f in fail:
        coach = f.replace("发送失败: 教练[", "").split("]")[0] if "教练[" in f else ""
        if coach:
            send_success[coach] = False
            if ", 错误: " in f:
                reason = f.split(", 错误: ")[-1]
                fail_reasons[coach] = reason
    
    # 获取根目录（脚本所在目录）
    root_dir = get_app_dir().resolve()
    log_path = write_csv_log(root_dir, timestamp, results, send_success, fail_reasons, warnings, errors)

    print("\n发送完成。")
    print(f"成功: {len(success)} 封, 失败: {len(fail)} 封")
    print(f"CSV日志：{log_path}")

    # 删除表格文件
    print("\n" + "=" * 70)
    clear_input = input("按回车键删除【学员数据表】和【学员对接表】（输入其他任意键跳过）...")
    if clear_input.strip() == "":
        student_deleted = delete_table_file(files["student"])
        mapping_deleted = delete_table_file(files["mapping"])
        
        print()
        if student_deleted and mapping_deleted:
            print("  表格文件已删除。")
        else:
            if not student_deleted:
                print(f"  学员数据表删除失败: {files['student'].name}")
            if not mapping_deleted:
                print(f"  学员对接表删除失败: {files['mapping'].name}")
    
    print("\n" + "=" * 70)
    input("按回车键退出...")


def build_parser():
    parser = argparse.ArgumentParser(description="教练学员信息邮件分发工具")
    parser.add_argument("--data-dir", default="", help="数据目录（包含三个表格和可选的sender_config.txt）")
    parser.add_argument("--sender-email", default="", help="发件人邮箱（命令行模式）")
    parser.add_argument("--auth-code", default="", help="SMTP授权码（命令行模式，可不填，运行时输入）")
    parser.add_argument("--smtp-host", default="", help="SMTP服务器地址")
    parser.add_argument("--smtp-port", type=int, default=465, help="SMTP SSL端口（默认: 465）")
    parser.add_argument("--sender-name", default="", help="发件人显示名称")
    parser.add_argument("--preview-only", action="store_true", help="仅生成预览与日志，不发送邮件")
    parser.add_argument("--interactive", "-i", action="store_true", help="强制使用交互式模式")
    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    try:
        run(args)
    except Exception as e:
        print(f"程序执行失败: {e}")
        # 双击运行时，让用户能看到错误信息
        if is_double_click_run():
            input("\n按回车键退出...")
        raise
