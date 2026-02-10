#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 BSC 计分规则解析修复的测试脚本
覆盖：全角字符、目标值百分比模式、扣减模式、按比例得分模式
"""

import sys
import math

# 将当前目录加入路径
sys.path.insert(0, '.')

from bsc_processor import BSCProcessor


def approx_eq(a, b, tol=0.01):
    """浮点近似比较"""
    return abs(a - b) < tol


def test_normalize_fullwidth():
    """测试全角→半角转换"""
    print("=== 测试 normalize_fullwidth ===")
    cases = [
        ('实际值＜85%＊目标值', '实际值<85%*目标值'),
        ('＞100%', '>100%'),
        ('＝80%', '=80%'),
        ('（注）', '(注)'),
        ('85％', '85%'),
    ]
    ok = True
    for inp, expected in cases:
        result = BSCProcessor.normalize_fullwidth(inp)
        passed = result == expected
        ok = ok and passed
        status = "PASS" if passed else "FAIL"
        print(f"  [{status}] '{inp}' -> '{result}' (expected '{expected}')")
    return ok


def test_extract_target_pct_baseline():
    """测试目标值百分比提取"""
    print("\n=== 测试 extract_target_pct_baseline ===")
    cases = [
        ("实际值<85%*目标值万元，不得分", 0.85, '正向'),
        ("实际值=80%*目标值万元，得60分", 0.80, '正向'),
        ("实际值<80%，不得分", 0.80, '正向'),
        ("<80%*目标值，不得分", 0.80, '正向'),
        ("实际值>120%*目标值，不得分", 1.20, '逆向'),
    ]
    ok = True
    for rule, expected_ratio, expected_dir in cases:
        # 先做全角转半角（与 calculate_baseline 一致）
        normalized = BSCProcessor.normalize_fullwidth(rule)
        result = BSCProcessor.extract_target_pct_baseline(normalized)
        if result is None:
            passed = False
            print(f"  [FAIL] '{rule}' -> None (expected ({expected_ratio}, '{expected_dir}'))")
        else:
            ratio, direction = result
            passed = approx_eq(ratio, expected_ratio) and direction == expected_dir
            status = "PASS" if passed else "FAIL"
            print(f"  [{status}] '{rule}' -> ({ratio}, '{direction}') (expected ({expected_ratio}, '{expected_dir}'))")
        ok = ok and passed
    return ok


def test_calculate_baseline_target_pct():
    """测试 calculate_baseline 对目标值百分比模式的完整计算"""
    print("\n=== 测试 calculate_baseline: 目标值百分比模式 ===")
    proc = BSCProcessor.__new__(BSCProcessor)

    cases = [
        # (rule_text, target, is_percent, expected_baseline, expected_status)
        ("实际值＜85%*目标值万元，不得分", 1594.49, False, 1355.32, '成功'),
        ("实际值＜80%*目标值万元，不得分", 83.27, False, 66.62, '成功'),
        ("实际值＜80%，不得分（按比例得分）", 126.0, False, 75.6, '成功'),
    ]
    ok = True
    for rule, target, is_pct, expected_bl, expected_st in cases:
        baseline, status, direction = proc.calculate_baseline(target, rule, is_pct)
        passed = approx_eq(baseline, expected_bl) and status == expected_st
        s = "PASS" if passed else "FAIL"
        print(f"  [{s}] rule='{rule}', target={target} -> baseline={baseline:.2f}, status='{status}' "
              f"(expected baseline={expected_bl}, status='{expected_st}')")
        ok = ok and passed
    return ok


def test_deduction_with_koujan():
    """测试扣减模式（扣减10分）"""
    print("\n=== 测试 extract_deduction_params: 扣减模式 ===")
    cases = [
        ("每高于1%，扣减10分", 1.0, 10.0, '逆向'),
        ("每高于目标值1%，扣减5分", 1.0, 5.0, '逆向'),
        ("每低于1%，扣减10分", 1.0, 10.0, '正向'),
    ]
    ok = True
    for rule, expected_x, expected_y, expected_dir in cases:
        normalized = BSCProcessor.normalize_fullwidth(rule)
        result = BSCProcessor.extract_deduction_params(normalized)
        if result is None:
            passed = False
            print(f"  [FAIL] '{rule}' -> None (expected match)")
        else:
            x, y, direction, has_pct = result
            passed = (approx_eq(x, expected_x) and approx_eq(y, expected_y)
                      and direction == expected_dir)
            s = "PASS" if passed else "FAIL"
            print(f"  [{s}] '{rule}' -> (x={x}, y={y}, dir='{direction}', pct={has_pct}) "
                  f"(expected x={expected_x}, y={expected_y}, dir='{expected_dir}')")
        ok = ok and passed
    return ok


def test_calculate_baseline_deduction_koujan():
    """测试 calculate_baseline 对扣减模式的完整计算"""
    print("\n=== 测试 calculate_baseline: 扣减模式 ===")
    proc = BSCProcessor.__new__(BSCProcessor)

    # 每高于1%，扣减10分 -> target=0.444, 允许偏差 = (40/10)*1% = 4% = 0.04
    # 逆向: baseline = 0.444 + 0.04 = 0.484
    baseline, status, direction = proc.calculate_baseline(0.444, "每高于1%，扣减10分", False)
    passed = approx_eq(baseline, 0.484) and status == '成功' and direction == '逆向'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '每高于1%，扣减10分', target=0.444 -> baseline={baseline:.4f}, "
          f"status='{status}', dir='{direction}' (expected 0.484, 成功, 逆向)")
    return passed


def test_ratio_baseline_extended():
    """测试扩展的比例型规则识别"""
    print("\n=== 测试 extract_ratio_baseline: 扩展模式 ===")
    cases = [
        ("按照达成比例得分，低于60分不得分", 0.6, '正向'),
        ("按实际完成比例得分，低于60分不得分", 0.6, '正向'),
        ("100分封顶，低于60分不得分", 0.6, '正向'),
    ]
    ok = True
    for rule, expected_ratio, expected_dir in cases:
        normalized = BSCProcessor.normalize_fullwidth(rule)
        result = BSCProcessor.extract_ratio_baseline(normalized)
        if result is None:
            passed = False
            print(f"  [FAIL] '{rule}' -> None (expected ({expected_ratio}, '{expected_dir}'))")
        else:
            ratio, direction = result
            passed = approx_eq(ratio, expected_ratio) and direction == expected_dir
            s = "PASS" if passed else "FAIL"
            print(f"  [{s}] '{rule}' -> ({ratio}, '{direction}') "
                  f"(expected ({expected_ratio}, '{expected_dir}'))")
        ok = ok and passed
    return ok


def test_calculate_baseline_ratio_extended():
    """测试 calculate_baseline 对按比例得分模式的完整计算"""
    print("\n=== 测试 calculate_baseline: 按比例得分 ===")
    proc = BSCProcessor.__new__(BSCProcessor)

    # 按照达成比例得分，低于60分不得分 -> target=0.85, ratio=0.6
    # baseline = 0.85 * 0.6 = 0.51
    baseline, status, direction = proc.calculate_baseline(
        0.85, "按照达成比例得分，低于60分不得分", False
    )
    passed = approx_eq(baseline, 0.51) and status == '成功'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '按照达成比例得分，低于60分不得分', target=0.85 -> "
          f"baseline={baseline:.4f}, status='{status}' (expected 0.51, 成功)")
    return passed


def test_regression_existing_patterns():
    """回归测试：确保原有模式不受影响"""
    print("\n=== 回归测试：原有模式 ===")
    proc = BSCProcessor.__new__(BSCProcessor)
    ok = True

    # 每低1%扣1分 (target=0.85) -> 允许偏差 = (40/1)*1% = 40% = 0.4
    # baseline = 0.85 - 0.4 = 0.45
    baseline, status, direction = proc.calculate_baseline(0.85, "每低1%扣1分", False)
    passed = approx_eq(baseline, 0.45) and status == '成功'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '每低1%扣1分', target=0.85 -> baseline={baseline:.4f}, "
          f"status='{status}' (expected 0.45, 成功)")
    ok = ok and passed

    # 完成3个得60分 -> baseline=3
    baseline, status, direction = proc.calculate_baseline(5, "完成3个，得60分", False)
    passed = approx_eq(baseline, 3.0) and status == '成功'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '完成3个，得60分', target=5 -> baseline={baseline:.4f}, "
          f"status='{status}' (expected 3.0, 成功)")
    ok = ok and passed

    # 低于85%不得分 -> baseline=0.85
    baseline, status, direction = proc.calculate_baseline(1.0, "低于85%不得分", True)
    passed = approx_eq(baseline, 0.85) and status == '成功'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '低于85%不得分', target=1.0 -> baseline={baseline:.4f}, "
          f"status='{status}' (expected 0.85, 成功)")
    ok = ok and passed

    return ok


def test_fullwidth_end_to_end():
    """端到端测试：全角字符规则 → calculate_baseline"""
    print("\n=== 端到端测试：全角字符 ===")
    proc = BSCProcessor.__new__(BSCProcessor)

    # 全角＜应被转换为半角<
    rule = "实际值＜85％＊目标值万元，不得分"
    baseline, status, direction = proc.calculate_baseline(1594.49, rule, False)
    passed = approx_eq(baseline, 1355.32) and status == '成功'
    s = "PASS" if passed else "FAIL"
    print(f"  [{s}] '{rule}', target=1594.49 -> baseline={baseline:.2f}, "
          f"status='{status}' (expected 1355.32, 成功)")
    return passed


def main():
    all_pass = True
    all_pass = test_normalize_fullwidth() and all_pass
    all_pass = test_extract_target_pct_baseline() and all_pass
    all_pass = test_calculate_baseline_target_pct() and all_pass
    all_pass = test_deduction_with_koujan() and all_pass
    all_pass = test_calculate_baseline_deduction_koujan() and all_pass
    all_pass = test_ratio_baseline_extended() and all_pass
    all_pass = test_calculate_baseline_ratio_extended() and all_pass
    all_pass = test_regression_existing_patterns() and all_pass
    all_pass = test_fullwidth_end_to_end() and all_pass

    print("\n" + "=" * 50)
    if all_pass:
        print("ALL TESTS PASSED")
    else:
        print("SOME TESTS FAILED")
    print("=" * 50)

    return 0 if all_pass else 1


if __name__ == '__main__':
    sys.exit(main())
