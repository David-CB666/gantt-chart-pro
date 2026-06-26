#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""非交互式甘特圖生成 wrapper — 用於 CI / Demo"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

# Monkey-patch input to return defaults
import builtins
_builtin_input = builtins.input
def _mock_input(prompt=""):
    prompt_lower = prompt.lower() if prompt else ""
    if "配色" in prompt_lower or "方案" in prompt_lower:
        return "0"   # default (blue_pro)
    if "日历" in prompt_lower or "日曆" in prompt_lower or "模式" in prompt_lower:
        return "calendar_days"
    if "后缀" in prompt_lower or "後綴" in prompt_lower:
        return ""
    if "退出" in prompt_lower or "enter" in prompt_lower:
        return ""
    if "路径" in prompt_lower or "路徑" in prompt_lower:
        return ""
    return ""
builtins.input = _mock_input

from gantt_chart_pro import create_gantt_pro

if __name__ == "__main__":
    input_file = sys.argv[1] if len(sys.argv) > 1 else "examples/EDF.xlsx"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "demo/demo_schedule.xlsx"
    print(f"Generating Gantt chart: {input_file} -> {output_file}")
    success = create_gantt_pro(input_file)
    if success:
        print(f"✓ Done: {output_file}")
    else:
        print("✗ Failed")
        sys.exit(1)
