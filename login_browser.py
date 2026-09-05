#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学习通账号登录（弹出浏览器，人工扫码/输密码，自动保存登录态）
================================================================

用法：
  python3 login_browser.py                 # 弹浏览器登录（会自动复用上次登录态）
  python3 login_browser.py --relogin       # 强制重新登录
  python3 login_browser.py --timeout 600   # 等待登录超时改为 600 秒

说明：
  · 脚本会弹出一个 Chromium 窗口并打开统一登录页，你在里面登录即可
    （学习通 App 扫码最省事；也可手机号+学习通密码、机构账号
    "中国科学院大学"+学号、短信验证码）
  · 登录成功后脚本自动检测并保存登录态（含 HttpOnly Cookie）到
    ~/.ucas_schedule/，之后直接运行 crawl.py 即可，无需重复登录
  · 依赖：pip3 install playwright && python3 -m playwright install chromium
"""

import argparse
import sys
from pathlib import Path

# 复用 skill 目录里的核心模块（单一数据源，避免复制代码）
sys.path.insert(0, str(Path(__file__).resolve().parent
                       / "ucas-schedule-export" / "scripts"))
import ucas_crawler  # noqa: E402


def main():
    ap = argparse.ArgumentParser(
        description="弹出浏览器登录学习通，自动保存登录态供 crawl.py 使用")
    ap.add_argument("--relogin", action="store_true", help="强制重新登录")
    ap.add_argument("--timeout", type=int, default=300, help="等待登录的秒数（默认300）")
    args = ap.parse_args()
    ucas_crawler.cmd_login_browser(args)


if __name__ == "__main__":
    main()
