#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
平衡计分卡KPI数据处理脚本
功能：将非结构化的KPI指标数据转化为标准化的平衡计分卡格式
"""

import pandas as pd
import re
import openpyxl
from pathlib import Path
from typing import Tuple, Dict, Optional, Union


class BSCProcessor:
    """平衡计分卡数据处理器"""

    def __init__(self, input_file: str):
        """
        初始化处理器

        Args:
            input_file: 输入Excel文件路径
        """
        self.input_file = input_file
        self.df = None
        self.target_col = None
        self.rule_col = None

    def load_data(self) -> pd.DataFrame:
        """读取Excel数据"""
        try:
            self.df = pd.read_excel(self.input_file)
            print(f"成功读取文件，共 {len(self.df)} 行，{len(self.df.columns)} 列")
            print(f"列名: {list(self.df.columns)}")
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
            print(f"识别到表头在第{header_row_idx + 1}行")
            self.df = pd.read_excel(self.input_file, header=header_row_idx)
            print(f"列名: {list(self.df.columns)}")

        # 识别目标值列（按优先级）
        self.target_col = None
        for keyword in target_keywords:
            for col in self.df.columns:
                col_str = str(col)
                # 匹配关键字，排除半年相关列
                if keyword in col_str and not is_excluded(col_str):
                    self.target_col = col
                    print(f"找到目标值列: {self.target_col} (匹配关键字: {keyword})")
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
                    print(f"找到计分规则列: {self.rule_col} (匹配关键字: {keyword})")
                    break
            if self.rule_col:
                break

        if not self.target_col:
            raise Exception("无法识别目标值列，请检查前3行是否有包含'目标值'或'全年目标值'的列")
        if not self.rule_col:
            raise Exception("无法识别计分规则列，请检查前3行是否有包含'计分规则'或'全年计分规则'的列")

        print(f"最终使用 - 目标值列: {self.target_col}")
        print(f"最终使用 - 计分规则列: {self.rule_col}")

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
            # 判断是否可能是百分比（0 < value < 1 且不是常见的整数）
            # 但这里我们保守处理，认为直接读入的数字就是实际值
            # 只有明确带%的才算百分比
            return float(value), False

        return 0.0, False

    @staticmethod
    def extract_explicit_baseline(rule_text: str) -> Optional[Tuple[float, str]]:
        """
        逻辑A：从规则文本中提取显式的底线值

        模式匹配：
        - "低于85%不得分"、"<85%不得分"、"<85%为0" -> 0.85, '正向'
        - "高于85%不得分"、">85%得0分" -> 0.85, '逆向'
        - "85分得60分"、"85为60分" -> 85, '正向'

        注意：如果匹配"低于80分不得分"这种得分阈值，应返回None让后续逻辑处理

        Args:
            rule_text: 规则文本

        Returns:
            (底线值, 指标方向) 或 None
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 先检查是否包含得分阈值描述（如"低于80分不得分"），这种情况下返回None
        # 因为这些值是得分阈值，不是指标底线值
        score_threshold_pattern = r'(?:低于|不足|少于)\s*([0-9]+)\s*分\s*(?:不得分|为0|得0分)'
        if re.search(score_threshold_pattern, rule_text):
            # 检查是否同时包含比例型关键词，如果是则让比例逻辑处理
            ratio_keywords = [
                r'实际\s*[:：]?\s*目标',
                r'达成\s*/\s*目标',
                r'除\s*以\s*目标',
                r'/\s*目标',
                r'÷\s*目标',
                r'最多\s*100分',
            ]
            if any(re.search(kw, rule_text) for kw in ratio_keywords):
                return None  # 让比例规则处理

        # 模式1: 低于/小于XX%不得分/得0分 (正向指标) - 必须带%才认为是指标值
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

        # 模式2: 高于/大于/超过XX%不得分/得0分 (逆向指标) - 必须带%才认为是指标值
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
        """
        提取比例型计分规则的得分阈值

        比例型规则：得分 = (实际值 ÷ 目标值) × 100
        - "实际达成率/目标值*100，低于60分不得分" -> 得分阈值60, 底线系数0.6
        - "最多100分，低于60分不得分" -> 得分阈值60, 底线系数0.6

        Args:
            rule_text: 规则文本

        Returns:
            (得分阈值, 方向) 或 None
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 检查是否是比例型规则（包含除法或公式形式）
        ratio_keywords = [
            r'实际\s*[:：]?\s*目标',
            r'达成\s*/\s*目标',
            r'除\s*以\s*目标',
            r'/\s*目标',
            r'÷\s*目标',
        ]

        is_ratio_rule = any(re.search(kw, rule_text) for kw in ratio_keywords)

        # 或者规则是"最多100分"这种形式（隐含比例计算）
        is_max_score = re.search(r'最多\s*100分', rule_text) is not None

        if not is_ratio_rule and not is_max_score:
            return None

        # 提取得分阈值
        # 模式：低于60分不得分、60分以下不得分、低于60为0、不足60分不得分等
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
                # 如果阈值是60，返回比例系数0.6
                # 其他阈值按比例计算
                ratio = threshold / 100
                return (ratio, '正向')

        # 默认使用60分
        return (0.6, '正向')

    @staticmethod
    def extract_deduction_params(rule_text: str) -> Optional[Tuple[float, float, str, bool]]:
        """
        从规则文本中提取扣分参数

        匹配模式：
        - "每低1%扣2分" -> (1, 2, '正向', True)
        - "每高1%扣2分" -> (1, 2, '逆向', True)
        - "每低于目标值0.1%扣5分" -> (0.1, 5, '正向', True)

        Args:
            rule_text: 规则文本

        Returns:
            (每X单位, 扣Y分, 方向, 规则中的X是否带%) 或 None
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text = rule_text.strip()

        # 正向指标：每低/每差/每降/每少/每低于目标值
        # 支持"扣"和"减"两种表达方式
        # 格式如：每低1%扣2分、每低于1%扣2分、每低于目标值2%，减5分、每低3分减7分、每少0.1%扣5分
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
            # 格式8: 每差/每降/每小 X单位[,，]扣/减Y分
            r'每[差小降](?:于目标值)?\s*([0-9]+\.?[0-9]*)\s*(?:%|[个人次项元万千百])?\s*[,，]?\s*(?:扣|减)\s*([0-9]+\.?[0-9]*)\s*分',
        ]

        for pattern in patterns_positive:
            match = re.search(pattern, rule_text)
            if match:
                x = float(match.group(1))
                y = float(match.group(2))
                # 检查匹配的文本中是否包含%
                has_percent = '%' in match.group(0)
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
                y = float(match.group(2))  # 扣分数是group(2)因为单位组是非捕获的
                # 检查匹配的文本中是否包含%
                has_percent = '%' in match.group(0)
                return (x, y, '逆向', has_percent)

        return None

    @staticmethod
    def detect_indicator_direction(rule_text: str) -> Optional[str]:
        """
        从规则文本中检测指标方向（正向/逆向）

        正向指标：越高越好（如完成率、达成率、收入）
        逆向指标：越低越好（如投诉率、差错率、成本）

        Args:
            rule_text: 规则文本

        Returns:
            '正向'/'逆向'/None
        """
        if pd.isna(rule_text) or not isinstance(rule_text, str):
            return None

        rule_text_lower = rule_text.lower()

        # 逆向指标关键词（优先检测，权重更高）
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
        """
        计算底线值（核心函数）

        逻辑优先级：
        A. 比例型规则：实际值/目标值×100，得分阈值60分
        B. 扣分推导：根据扣分规则计算
        C. 显式阈值：从规则中直接提取（排除得分阈值如"60分"）
        D. 默认兜底：目标值的80%或120%

        Args:
            target: 目标值（已归一化为浮点数）
            rule_text: 规则文本
            is_percent: 原始数据是否为百分比格式

        Returns:
            (底线值, 解析状态, 指标方向)
        """
        baseline = None
        status = '成功'
        direction = self.detect_indicator_direction(rule_text) or '正向'

        # 逻辑A（最高优先级）: 比例型规则
        # 得分 = (实际值 ÷ 目标值) × 100
        # 当得分=60时：底线值 = 目标值 × 0.6
        ratio_info = self.extract_ratio_baseline(rule_text)
        if ratio_info is not None:
            ratio, detected_direction = ratio_info
            direction = detected_direction
            baseline = target * ratio
            return baseline, '成功', direction

        # 逻辑B: 扣分推导
        deduction = self.extract_deduction_params(rule_text)
        if deduction is not None:
            x, y, deduction_direction, rule_has_percent = deduction
            direction = deduction_direction

            # 假设满分100，底线60分，允许扣40分
            allowed_gap = (40 / y) * x

            # 处理百分比单位转换
            if rule_has_percent:
                allowed_gap = allowed_gap / 100

            if direction == '逆向':
                baseline = target + allowed_gap
            else:
                baseline = target - allowed_gap

            return baseline, '成功', direction

        # 逻辑B: 显式阈值（排除得分阈值如"60分"、"不得分"）
        explicit_baseline = self.extract_explicit_baseline(rule_text)
        if explicit_baseline is not None:
            baseline, detected_direction = explicit_baseline
            # 检查是否是得分阈值（如60分这种常见的得分底线）
            # 如果提取的值是60且规则中包含"不得分/得0分"，可能是得分阈值而非指标值
            is_score_threshold = (
                abs(baseline - 60) < 0.01 and  # 值接近60
                '不得分' in rule_text          # 包含"不得分"
            )
            # 如果是得分阈值且有扣分规则，则不应使用显式阈值
            # 但因为已经检查过扣分规则，这里能执行说明没有扣分规则
            # 所以只有当baseline看起来像指标值时才使用
            if not is_score_threshold:
                if detected_direction:
                    direction = detected_direction
                return baseline, '成功', direction

            return baseline, '成功', direction

        # 逻辑C: 默认兜底
        if direction == '逆向':
            baseline = target * 1.2
        else:
            baseline = target * 0.8

        status = '人工校验'
        return baseline, status, direction

    @staticmethod
    def format_value(value: float, is_percent: bool) -> str:
        """
        将数值格式化为显示字符串

        Args:
            value: 数值
            is_percent: 是否为百分比格式

        Returns:
            格式化后的字符串
        """
        if is_percent:
            # 转换为百分比显示
            return f"{value * 100:.2f}%"
        else:
            # 根据数值大小决定小数位数
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
        """
        生成规范化计分规则文案

        Args:
            target: 目标值
            baseline: 底线值
            direction: 指标方向 ('正向'/'逆向')
            is_percent: 是否为百分比格式

        Returns:
            规范化文案
        """
        target_str = self.format_value(target, is_percent)
        baseline_str = self.format_value(baseline, is_percent)

        if direction == '逆向':
            # 逆向指标
            template = (
                f"P为指标实际值，{target_str}为目标值，{baseline_str}为底线值。\n"
                f"1.若P≤{target_str}，得100分（满分）；\n"
                f"2.若{target_str}<P<{baseline_str}，按线性比例计算，即："
                f"得分=100-(P-{target_str})/({baseline_str}-{target_str})×(100-60)；\n"
                f"3.若P={baseline_str}，得60分（基础分）；\n"
                f"4.若P＞{baseline_str}，得0分。"
            )
        else:
            # 正向指标
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
        """
        处理单行数据

        Args:
            row: 数据行

        Returns:
            包含处理结果的字典
        """
        target_raw = row[self.target_col]
        rule_text = row[self.rule_col] if self.rule_col in row else ""

        # 归一化目标值
        target, is_percent = self.normalize_target_value(target_raw)

        # 计算底线值
        baseline, status, direction = self.calculate_baseline(
            target, str(rule_text), is_percent
        )

        # 根据底线值和目标值的关系，验证/调整方向（仅在Manual Check或未明确检测到方向时）
        # 对于显式提取的方向（状态为Success），保持原方向不变
        if status == '人工校验':
            if baseline < target:
                direction = '正向'
            elif baseline > target:
                direction = '逆向'

        # 生成规范化文案
        standard_rule = self.generate_standard_rule(
            target, baseline, direction, is_percent
        )

        # 格式化底线值用于显示
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

    def process(self) -> pd.DataFrame:
        """
        执行完整处理流程

        Returns:
            处理后的DataFrame
        """
        # 加载数据
        self.load_data()

        # 识别列
        self.identify_columns()

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
                print(f"已过滤掉 {filtered_rows} 行半年指标数据")
            self.df = filtered_df
        else:
            print("警告：未找到可用于过滤的列，跳过半年指标过滤")

        # 添加新列
        results = {
            '推导底线值': [],
            '规范版计分规则': [],
            '解析状态': [],
            '指标方向': []
        }

        # 处理每一行
        for idx, row in self.df.iterrows():
            try:
                result = self.process_row(row)
                results['推导底线值'].append(result['推导底线值'])
                results['规范版计分规则'].append(result['规范版计分规则'])
                results['解析状态'].append(result['解析状态'])
                results['指标方向'].append(result['指标方向'])
            except Exception as e:
                # 异常处理：使用默认值
                print(f"警告: 第{idx+2}行处理失败: {e}")
                results['推导底线值'].append('ERROR')
                results['规范版计分规则'].append(f'解析失败: {str(e)}')
                results['解析状态'].append(f'ERROR: {str(e)[:50]}')
                results['指标方向'].append('unknown')

        # 将结果添加到原DataFrame
        for col, values in results.items():
            self.df[col] = values

        # 统计
        status_counts = self.df['解析状态'].value_counts()
        print(f"\n处理完成！统计信息：")
        print(status_counts)

        manual_check = (self.df['解析状态'] == '人工校验').sum()
        print(f"需要人工核对的行数: {manual_check}")

        return self.df

    def save(self, output_file: str = 'processed_scorecard.xlsx'):
        """
        保存处理结果

        Args:
            output_file: 输出文件名
        """
        if self.df is None:
            raise Exception("请先执行process()方法")

        # 设置列宽
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            self.df.to_excel(writer, index=False)

            # 获取workbook和worksheet
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
                    # 找到列的字母索引
                    col_idx = list(self.df.columns).index(col) + 1
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    worksheet.column_dimensions[col_letter].width = width

            # 设置规范版计分规则列为自动换行
            if '规范版计分规则' in self.df.columns:
                col_idx = list(self.df.columns).index('规范版计分规则') + 1
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                for cell in worksheet[col_letter]:
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

        print(f"\n结果已保存到: {output_file}")
        return output_file


def main():
    """主函数"""
    import sys

    # 获取输入文件
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # 查找当前目录下的Excel文件
        current_dir = Path(__file__).parent
        excel_files = list(current_dir.glob('*.xlsx')) + list(current_dir.glob('*.xls'))
        if not excel_files:
            print("当前目录下没有找到Excel文件，请指定输入文件路径")
            print("用法: python bsc_processor.py <input_file>")
            return

        # 选择第一个Excel文件
        input_file = str(excel_files[0])
        print(f"自动选择输入文件: {input_file}")

    try:
        # 创建处理器
        processor = BSCProcessor(input_file)

        # 执行处理
        processor.process()

        # 保存结果
        processor.save()

        print("\n处理完成！")

    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
