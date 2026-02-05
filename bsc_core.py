#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
平衡计分卡KPI数据处理核心类
功能：将非结构化的KPI指标数据转化为标准化的平衡计分卡格式
"""

import pandas as pd
import re
import openpyxl
from typing import Tuple, Dict, Optional, Union
from io import BytesIO


class BSCProcessor:
    """平衡计分卡数据处理器核心类"""

    def __init__(self, input_file=None):
        """
        初始化处理器

        Args:
            input_file: 输入Excel文件路径或BytesIO对象
        """
        self.input_file = input_file
        self.df = None
        self.target_col = None
        self.rule_col = None
        self.status_messages = []  # 用于存储处理状态信息

    def _log(self, message: str):
        """记录日志信息"""
        self.status_messages.append(message)

    def load_data(self) -> pd.DataFrame:
        """读取Excel数据"""
        try:
            self.df = pd.read_excel(self.input_file)
            self._log(f"成功读取文件，共 {len(self.df)} 行，{len(self.df.columns)} 列")
            self._log(f"列名: {list(self.df.columns)}")
            return self.df
        except Exception as e:
            raise Exception(f"读取文件失败: {e}")

    def identify_columns(self) -> Tuple[str, str]:
        """
        自动识别目标值列和计分规则列
        优先级：全年 > 年度 > 其他，排除所有半年相关列

        Returns:
            (目标值列名, 计分规则列名)
        """
        # 目标值关键字（按优先级排序：全年 > 年度 > 其他）
        target_keywords = [
            '全年目标值',
            '年度目标值', '2026目标值', '26年度目标值',
            '目标值', '考核目标值', 'kpi目标值', '指标值', '目标'
        ]
        # 计分规则关键字（按优先级排序：全年 > 年度 > 其他）
        rule_keywords = [
            '全年计分规则',
            '计分规则', '年度计分规则',
            '评分规则', '考核规则', '计分标准'
        ]

        # 需要排除的关键词（半年相关）
        exclude_keywords = ['半年度', '半年', '半期', '中期']

        def is_excluded(col_str: str) -> bool:
            """检查列名是否包含排除关键词"""
            return any(kw in col_str for kw in exclude_keywords)

        # 检查表头在哪一行（最多检查前3行）
        header_row_idx = None
        max_check_rows = min(4, len(self.df) + 1)

        # 首先检查当前列名（第0行，作为表头）
        for keyword in target_keywords:
            for col in self.df.columns:
                col_str = str(col)
                # 匹配关键字，但排除半年相关列
                if keyword in col_str and not is_excluded(col_str):
                    if header_row_idx is None:
                        header_row_idx = 0
                    break
            if header_row_idx == 0:
                break

        for keyword in rule_keywords:
            for col in self.df.columns:
                col_str = str(col)
                # 匹配关键字，但排除半年相关列
                if keyword in col_str and not is_excluded(col_str):
                    if header_row_idx is None:
                        header_row_idx = 0
                    break
            if header_row_idx == 0:
                break

        # 如果第一行列名没找到所有关键字，检查数据行
        if header_row_idx is None or True:
            found_target_in_col = any(
                any(kw in str(col) and not is_excluded(str(col)) for kw in target_keywords)
                for col in self.df.columns
            )
            found_rule_in_col = any(
                any(kw in str(col) and not is_excluded(str(col)) for kw in rule_keywords)
                for col in self.df.columns
            )

            for i in range(min(3, len(self.df))):
                row_data = self.df.iloc[i].astype(str)
                has_target = any(
                    any(kw in str(val) for kw in target_keywords)
                    for val in row_data.values
                )
                has_rule = any(
                    any(kw in str(val) for kw in rule_keywords)
                    for val in row_data.values
                )

                if has_target and has_rule:
                    header_row_idx = i + 1
                    break
                if has_target and not found_target_in_col:
                    header_row_idx = i + 1
                    found_target_in_col = True
                if has_rule and not found_rule_in_col:
                    if header_row_idx is None:
                        header_row_idx = i + 1
                    found_rule_in_col = True

        # 如果表头不在第0行，重新读取数据
        if header_row_idx and header_row_idx > 0:
            self._log(f"识别到表头在第{header_row_idx + 1}行")
            self.df = pd.read_excel(self.input_file, header=header_row_idx)
            self._log(f"列名: {list(self.df.columns)}")

        # 识别目标值列（按优先级）
        self.target_col = None
        for keyword in target_keywords:
            for col in self.df.columns:
                col_str = str(col)
                # 匹配关键字，排除半年相关列
                if keyword in col_str and not is_excluded(col_str):
                    self.target_col = col
                    self._log(f"找到目标值列: {self.target_col} (匹配关键字: {keyword})")
                    break
            if self.target_col:
                break

        # 识别计分规则列（按优先级）
        self.rule_col = None
        for keyword in rule_keywords:
            for col in self.df.columns:
                col_str = str(col)
                # 匹配关键字，排除半年相关列
                if keyword in col_str and not is_excluded(col_str):
                    self.rule_col = col
                    self._log(f"找到计分规则列: {self.rule_col} (匹配关键字: {keyword})")
                    break
            if self.rule_col:
                break

        if not self.target_col:
            raise Exception("无法识别目标值列，请检查前3行是否有包含'目标值'或'全年目标值'的列")
        if not self.rule_col:
            raise Exception("无法识别计分规则列，请检查前3行是否有包含'计分规则'或'全年计分规则'的列")

        self._log(f"最终使用 - 目标值列: {self.target_col}")
        self._log(f"最终使用 - 计分规则列: {self.rule_col}")

        return self.target_col, self.rule_col

    @staticmethod
    def normalize_target_value(value: Union[str, float, int]) -> Tuple[float, bool]:
        """
        归一化目标值
        处理多种情况：
        1. 字符串 "85%" -> 0.85, is_percent=True
        2. 字符串 "85分" -> 85, is_percent=False
        3. 数字 0.85 或 90 -> 对应浮点数, is_percent=False

        Args:
            value: 原始值

        Returns:
            (标准化后的浮点数, 是否为百分比格式)
        """
        if pd.isna(value):
            return 0.0, False

        # 如果是字符串
        if isinstance(value, str):
            value = value.strip()
            # 去除逗号（如 1,000）
            value = value.replace(',', '')

            # 检查是否包含百分号
            if '%' in value:
                num_str = value.replace('%', '').strip()
                try:
                    return float(num_str) / 100, True
                except ValueError:
                    return 0.0, False

            # 去除"分"字（如 "85分" -> 85）
            value = value.replace('分', '').strip()

            # 尝试直接转换
            try:
                return float(value), False
            except ValueError:
                return 0.0, False

        # 如果是数字
        elif isinstance(value, (int, float)):
            return float(value), False

        return 0.0, False

    @staticmethod
    def extract_explicit_baseline(rule_text: str) -> Optional[Tuple[float, str]]:
        """
        逻辑A：从规则文本中提取显式的底线值
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 先检查是否包含得分阈值描述
        score_threshold_pattern = r'(?:低于|不足|少于)\s*([0-9]+)\s*分\s*(?:不得分|为0|得0分)'
        if re.search(score_threshold_pattern, rule_text):
            ratio_keywords = [
                r'实际\s*[:：]?\s*目标',
                r'达成\s*/\s*目标',
                r'除\s*以\s*目标',
                r'/\s*目标',
                r'÷\s*目标',
                r'最多\s*100分',
            ]
            if any(re.search(kw, rule_text) for kw in ratio_keywords):
                return None

        # 模式1: 低于/小于XX%不得分/得0分 (正向指标)
        patterns_positive = [
            r'(?:低于|小于)\s*([0-9]+\.?[0-9]*)%\s*(?:不得分|得0分)',
            r'<\s*([0-9]+\.?[0-9]*)%\s*(?:不得分|得0分)',
            r'(?:低于|小于)\s*([0-9]+\.?[0-9]*)%\s*分?\s*(?:不得分|得0分)',
            r'<\s*([0-9]+\.?[0-9]*)%\s*分?\s*(?:不得分|得0分)',
        ]

        for pattern in patterns_positive:
            match = re.search(pattern, rule_text)
            if match:
                value = float(match.group(1))
                if '%' in match.group(0):
                    return value / 100, '正向'
                return value, '正向'

        # 模式2: 高于/大于/超过XX%不得分/得0分 (逆向指标)
        patterns_negative = [
            r'(?:高于|大于|超过)\s*([0-9]+\.?[0-9]*)%\s*(?:不得分|得0分)',
            r'>\s*([0-9]+\.?[0-9]*)%\s*(?:不得分|得0分)',
            r'(?:高于|大于|超过)\s*([0-9]+\.?[0-9]*)%\s*分?\s*(?:不得分|得0分)',
            r'>\s*([0-9]+\.?[0-9]*)%\s*分?\s*(?:不得分|得0分)',
        ]

        for pattern in patterns_negative:
            match = re.search(pattern, rule_text)
            if match:
                value = float(match.group(1))
                if '%' in match.group(0):
                    return value / 100, '逆向'
                return value, '逆向'

        # 模式3: XX分得60分/为60分
        patterns_60 = [
            r'([0-9]+\.?[0-9]*)%?\s*得60分',
            r'([0-9]+\.?[0-9]*)%?\s*[是为]60分',
            r'60分[是为]\s*([0-9]+\.?[0-9]*)%?',
        ]

        for pattern in patterns_60:
            match = re.search(pattern, rule_text)
            if match:
                value = float(match.group(1))
                if '%' in match.group(0):
                    return value / 100, '正向'
                return value, '正向'

        return None

    @staticmethod
    def extract_ratio_baseline(rule_text: str) -> Optional[Tuple[float, str]]:
        """提取比例型计分规则的得分阈值"""
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 检查是否是比例型规则
        ratio_keywords = [
            r'实际\s*[:：]?\s*目标',
            r'达成\s*/\s*目标',
            r'除\s*以\s*目标',
            r'/\s*目标',
            r'÷\s*目标',
        ]

        is_ratio_rule = any(re.search(kw, rule_text) for kw in ratio_keywords)
        is_max_score = re.search(r'最多\s*100分', rule_text) is not None

        if not is_ratio_rule and not is_max_score:
            return None

        # 提取得分阈值
        score_threshold_patterns = [
            r'低于\s*([0-9]+)\s*分[,，]?\s*(?:不得分|为0|得0分)',
            r'(?:不足|少于)\s*([0-9]+)\s*分[,，]?\s*(?:不得分|为0|得0分)',
            r'([0-9]+)分\s*(?:以下|为0)\s*(?:不得分|得0分)',
            r'满\s*100分.*?(?:低于|不足)\s*([0-9]+)\s*分[,，]?\s*不得分',
        ]

        for pattern in score_threshold_patterns:
            match = re.search(pattern, rule_text)
            if match:
                threshold = float(match.group(1))
                ratio = threshold / 100
                return (ratio, '正向')

        return (0.6, '正向')

    @staticmethod
    def extract_deduction_params(rule_text: str) -> Optional[Tuple[float, float, str, bool]]:
        """从规则文本中提取扣分参数"""
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 正向指标：每低/每差/每降/每少/每低于目标值/每起
        # 支持"扣"和"减"两种表达方式
        # 格式如：每低1%扣2分、每低于1%扣2分、每低于目标值2%，减5分、每低3分减7分、每少0.1%扣5分、每起扣10分
        # 注意：alternation要按长度降序排列（长的在前）避免部分匹配问题
        patterns_positive = [
            # 格式1: 每低于目标值X%[,，]扣/减Y分
            r'每低于目标值\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式2: 每低于X%[,，]扣/减Y分
            r'每低于\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式3: 每低X%[,，]扣/减Y分
            r'每低\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式4: 每少X%[,，]扣/减Y分（如：每少0.1%扣5分）
            r'每少\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式5: 每低于目标值X分[,，]扣/减Y分
            r'每低于目标值\s*([0-9]+\.?[0-9]*)\s*分\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式6: 每低X分[,，]扣/减Y分
            r'每低\s*([0-9]+\.?[0-9]*)\s*分\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式7: 每少X个[,，]扣/减Y分
            r'每少\s*([0-9]+\.?[0-9]*)\s*个\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式8: 每起X扣/减Y分 或 每起扣/减Y分（逆向指标：如"每起扣10分"表示事故类指标）
            # 这个格式需要特殊处理，因为默认是逆向指标（越少越好）
            r'每起\s*(?:([0-9]+\.?[0-9]*)\s*(?:个|单位|起)?\s*)?(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式9: 每X起扣/减Y分（如：每1起扣10分，逆向指标）
            r'每\s*([0-9]+\.?[0-9]*)\s*起\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            # 格式9: 每差/每降/每小 X单位[,，]扣/减Y分
            r'每[差小降](?:于目标值)?\s*([0-9]+\.?[0-9]*)\s*(?:%|[个人次项元万千百])?\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
        ]

        for pattern in patterns_positive:
            match = re.search(pattern, rule_text)
            if match:
                x_str = match.group(1)
                y_str = match.group(2)
                # 处理x为空的情况（如"每起扣10分"，默认每起算1个单位）
                if not x_str or x_str == '':
                    x = 1.0
                    y = float(y_str)
                else:
                    x = float(x_str)
                    y = float(y_str)
                # 检查匹配的文本中是否包含%
                matched_text = match.group(0)
                has_percent = '%' in matched_text
                return (x, y, '正向', has_percent)

        # 逆向指标：每高/每超/每多/每高于目标值
        patterns_negative = [
            r'每高于目标值\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            r'每高\s*([0-9]+\.?[0-9]*)%\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            r'每高于目标值\s*([0-9]+\.?[0-9]*)\s*分\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            r'每高\s*([0-9]+\.?[0-9]*)\s*分\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            r'每[高超多](?:于目标值)?\s*([0-9]+\.?[0-9]*)\s*(?:%|[个人次项元万千百])?\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
        ]

        for pattern in patterns_negative:
            match = re.search(pattern, rule_text)
            if match:
                x = float(match.group(1))
                y = float(match.group(2))
                matched_text = match.group(0)
                has_percent = '%' in matched_text
                return (x, y, '逆向', has_percent)

        return None

    @staticmethod
    def extract_accident_params(rule_text: str) -> Optional[Tuple[float, float, str]]:
        """
        提取"每起扣X分"类型的参数（逆向指标：事故类）

        这类指标的逻辑是：从0开始（满分），每有1起事故扣X分
        底线值（得60分）= 允许扣40分时的事故数 = 40 / X

        Args:
            rule_text: 规则文本

        Returns:
            (每起扣X分, 方向) 或 None
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 匹配"每起扣X分"或"每X起扣Y分"
        patterns = [
            r'每起\s*(?:([0-9]+\.?[0-9]*)\s*(?:个|单位|起)?\s*)?(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
            r'每\s*([0-9]+\.?[0-9]*)\s*起\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
        ]

        for pattern in patterns:
            match = re.search(pattern, rule_text)
            if match:
                x_str = match.group(1)  # 每X起（可能为空）
                y_str = match.group(2) if len(match.groups()) >= 2 else match.group(1)  # 扣Y分
                # 如果x_str为空，说明是"每起扣Y分"，默认每1起扣Y分
                if not x_str or x_str == '':
                    y = float(y_str)
                    x = 1.0
                else:
                    x = float(x_str)
                    y = float(y_str)
                return (x, y, '逆向')

        return None

    @staticmethod
    def detect_indicator_direction(rule_text: str) -> Optional[str]:
        """从规则文本中检测指标方向（正向/逆向）"""
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        # 逆向指标关键词
        negative_keywords = [
            '投诉率', '差错率', '故障率', '缺陷率', '不良率', '报废率',
            '成本控制', '控制在', '不超过', '不高于', '每高', '每超',
            '下降', '降低', '越低', '超出', '超支',
        ]

        # 正向指标关键词
        positive_keywords = [
            '完成率', '达成率', '实现率', '增长率', '提升率',
            '收入', '利润', '销售额', '产量', '达标', '超额',
            '越高', '不少于', '每低', '每降', '每差', '每少',
        ]

        positive_score = sum(1 for kw in positive_keywords if kw in rule_text)
        negative_score = sum(1 for kw in negative_keywords if kw in rule_text)

        if positive_score > negative_score:
            return '正向'
        elif negative_score > positive_score:
            return '逆向'
        return None

    def calculate_baseline(self, target: float, rule_text: str,
                          is_percent: bool) -> Tuple[float, str, str]:
        """计算底线值（核心函数）"""
        baseline = None
        status = '成功'
        direction = self.detect_indicator_direction(rule_text) or '正向'

        # 逻辑A（最高优先级）: 事故类指标（每起扣X分，逆向）
        # 这类指标从0开始（满分），每有1起事故扣X分
        # 底线值（得60分）= 允许扣40分时的事故数 = 40 / X
        accident_params = self.extract_accident_params(rule_text)
        if accident_params is not None:
            x, y, accident_params_direction = accident_params
            direction = accident_params_direction
            # 从0开始，每X起扣Y分，得60分时可以有多少起？
            # 40分 / Y分 = 允许扣的次数
            # 每次扣X起，所以允许起数 = (40 / Y) × X
            allowed_count = (40 / y) * x
            baseline = allowed_count
            return baseline, '成功', direction

        # 逻辑B: 比例型规则
        ratio_info = self.extract_ratio_baseline(rule_text)
        if ratio_info is not None:
            ratio, detected_direction = ratio_info
            direction = detected_direction
            baseline = target * ratio
            return baseline, '成功', direction

        # 逻辑C: 扣分推导
        deduction = self.extract_deduction_params(rule_text)
        if deduction is not None:
            x, y, deduction_direction, rule_has_percent = deduction
            direction = deduction_direction

            allowed_gap = (40 / y) * x

            if rule_has_percent:
                allowed_gap = allowed_gap / 100

            if direction == '逆向':
                baseline = target + allowed_gap
            else:
                baseline = target - allowed_gap

            return baseline, '成功', direction

        # 逻辑C: 显式阈值
        explicit_baseline = self.extract_explicit_baseline(rule_text)
        if explicit_baseline is not None:
            baseline, detected_direction = explicit_baseline
            if not (
                abs(baseline - 60) < 0.01 and '不得分' in rule_text
            ):
                if detected_direction:
                    direction = detected_direction
                return baseline, '成功', direction

            return baseline, '成功', direction

        # 逻辑D: 默认兜底
        if direction == '逆向':
            baseline = target * 1.2
        else:
            baseline = target * 0.8

        status = '人工校验'
        return baseline, status, direction

    @staticmethod
    def format_value(value: float, is_percent: bool) -> str:
        """将数值格式化为显示字符串"""
        if is_percent:
            return f"{value * 100:.2f}%"
        else:
            if value == int(value):
                return str(int(value))
            elif abs(value) >= 1000:
                return f"{value:.2f}"
            elif abs(value) >= 100:
                return f"{value:.2f}"
            else:
                return f"{value:.4f}".rstrip('0').rstrip('.')

    def generate_standard_rule(self, target: float, baseline: float,
                               direction: str, is_percent: bool) -> str:
        """生成规范化计分规则文案"""
        target_str = self.format_value(target, is_percent)
        baseline_str = self.format_value(baseline, is_percent)

        if direction == '逆向':
            template = (
                f"P为指标实际值，{target_str}为目标值，{baseline_str}为底线值。\n"
                f"1.若P≤{target_str}，得100分（满分）；\n"
                f"2.若{target_str}<P<{baseline_str}，按线性比例计算，即："
                f"得分=100-(P-{target_str})/({baseline_str}-{target_str})×(100-60)；\n"
                f"3.若P={baseline_str}，得60分（基础分）；\n"
                f"4.若P＞{baseline_str}，得0分。"
            )
        else:
            template = (
                f"P为指标实际值，{target_str}为目标值，{baseline_str}为底线值。\n"
                f"1.若P≥{target_str}，得100分（满分）；\n"
                f"2.若{baseline_str}<P<{target_str}，按线性比例计算，即："
                f"得分=60+(P-{baseline_str})/({target_str}-{baseline_str})×(100-60)；\n"
                f"3.若P={baseline_str}，得60分（基础分）；\n"
                f"4.若P<{baseline_str}，得0分。"
            )

        return template

    def process_row(self, row: pd.Series) -> Dict:
        """处理单行数据"""
        target_raw = row[self.target_col]
        rule_text = row[self.rule_col] if self.rule_col in row else ""

        target, is_percent = self.normalize_target_value(target_raw)
        baseline, status, direction = self.calculate_baseline(
            target, str(rule_text), is_percent
        )

        if status == '人工校验':
            if baseline < target:
                direction = '正向'
            elif baseline > target:
                direction = '逆向'

        standard_rule = self.generate_standard_rule(
            target, baseline, direction, is_percent
        )

        baseline_display = self.format_value(baseline, is_percent)

        return {
            '推导底线值': baseline_display,
            '规范版计分规则': standard_rule,
            '解析状态': status,
            '指标方向': direction,
            '目标值_归一化': target,
            '底线值_归一化': baseline,
            '是否百分比': is_percent
        }

    def process(self, progress_callback=None) -> pd.DataFrame:
        """
        执行完整处理流程

        Args:
            progress_callback: 进度回调函数，参数为当前进度(0-100)

        Returns:
            处理后的DataFrame
        """
        self.status_messages.clear()

        # 加载数据
        self.load_data()
        if progress_callback:
            progress_callback(20)

        # 识别列
        self.identify_columns()
        if progress_callback:
            progress_callback(40)

        # 过滤掉半年指标行
        original_rows = len(self.df)
        exclude_keywords = ['半年度', '半年', '半期', '中期']

        # 检查是否有"维度"或"评价指标"列用于判断
        filter_col = None
        for col in self.df.columns:
            if '维度' in str(col) or '评价指标' in str(col) or '指标名称' in str(col):
                filter_col = col
                break

        if filter_col:
            # 过滤包含半年关键词的行
            mask = self.df[filter_col].astype(str).apply(
                lambda x: not any(kw in x for kw in exclude_keywords)
            )
            filtered_df = self.df[mask].copy()
            filtered_rows = original_rows - len(filtered_df)
            if filtered_rows > 0:
                self._log(f"已过滤掉 {filtered_rows} 行半年指标数据")
            self.df = filtered_df
        else:
            self._log("警告：未找到可用于过滤的列，跳过半年指标过滤")

        # 添加新列
        results = {
            '推导底线值': [],
            '规范版计分规则': [],
            '解析状态': [],
            '指标方向': []
        }

        total_rows = len(self.df)
        for idx, row in self.df.iterrows():
            try:
                result = self.process_row(row)
                results['推导底线值'].append(result['推导底线值'])
                results['规范版计分规则'].append(result['规范版计分规则'])
                results['解析状态'].append(result['解析状态'])
                results['指标方向'].append(result['指标方向'])
            except Exception as e:
                self._log(f"警告: 第{idx+2}行处理失败: {e}")
                results['推导底线值'].append('ERROR')
                results['规范版计分规则'].append(f'解析失败: {str(e)}')
                results['解析状态'].append(f'ERROR: {str(e)[:50]}')
                results['指标方向'].append('unknown')

            # 更新进度
            if progress_callback:
                progress = 40 + int((idx + 1) / total_rows * 50)
                progress_callback(progress)

        # 将结果添加到原DataFrame
        for col, values in results.items():
            self.df[col] = values

        # 统计
        status_counts = self.df['解析状态'].value_counts()
        self._log("处理完成！统计信息：")
        for status, count in status_counts.items():
            self._log(f"  {status}: {count} 行")

        manual_check = (self.df['解析状态'] == '人工校验').sum()
        self._log(f"需要人工核对的行数: {manual_check}")

        if progress_callback:
            progress_callback(100)

        return self.df

    def save_to_bytesio(self) -> BytesIO:
        """
        将处理结果保存到BytesIO对象（用于Streamlit下载）

        Returns:
            BytesIO对象
        """
        if self.df is None:
            raise Exception("请先执行process()方法")

        output = BytesIO()

        # 设置列宽
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.df.to_excel(writer, index=False)

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # 设置列宽
            column_widths = {
                '推导底线值': 15,
                '规范版计分规则': 60,
                '解析状态': 20,
                '指标方向': 12
            }

            for col, width in column_widths.items():
                if col in self.df.columns:
                    col_idx = list(self.df.columns).index(col) + 1
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    worksheet.column_dimensions[col_letter].width = width

            # 设置规范版计分规则列为自动换行
            if '规范版计分规则' in self.df.columns:
                col_idx = list(self.df.columns).index('规范版计分规则') + 1
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                for cell in worksheet[col_letter]:
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

        output.seek(0)
        return output

    def get_stats(self) -> Dict:
        """获取处理统计信息"""
        if self.df is None or '解析状态' not in self.df.columns:
            return {}

        status_counts = self.df['解析状态'].value_counts()
        return {
            'total': len(self.df),
            'success': int(status_counts.get('成功', 0)),
            'manual_check': int(status_counts.get('人工校验', 0)),
            'error': int(sum(cnt for status, cnt in status_counts.items() if 'ERROR' in status)),
        }

    def get_logs(self) -> list:
        """获取处理日志"""
        return self.status_messages
