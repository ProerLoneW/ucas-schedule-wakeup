#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
爬取国科大全部学期课表，导出为 WakeUp 课程表可导入的 Excel/CSV
================================================================

登录态来源（按优先级，任选其一即可）：
  1. 环境变量 UCAS_COOKIE（Cookie 字符串，或写入同目录 .env 文件）
  2. 之前运行 login_browser.py / save-cookie 保存的登录态（~/.ucas_schedule/）

用法：
  # 方式一：环境变量传入 Cookie（推荐写进 .env，见 .env.example）
  export UCAS_COOKIE="route=...; UID=...; ..."
  python3 crawl.py

  # 方式二：先用 login_browser.py 登录一次，之后直接
  python3 crawl.py

  # 常用参数
  python3 crawl.py --out ./output        # 指定输出目录（默认当前目录）
  python3 crawl.py --no-merge            # 教务课程不合并，导出更细粒度记录
  python3 crawl.py --cookie "..."        # 临时用这串 Cookie（不写文件）

输出（每个有课的学期一对文件 + 完整原始数据）：
  WakeUp课表_<学年>第<学期>学期.xlsx / .csv   —— 7 列 WakeUp 导入格式
  schedule_all.json                          —— 接口返回的完整原始数据

Cookie 怎么拿：浏览器登录课表页后，F12 -> Network -> 刷新 -> 点任意
kb.mooc.ucas.edu.cn 请求 -> Request Headers -> 复制整串 Cookie 值。
注意必须从请求头复制（含 HttpOnly 凭据），控制台 document.cookie 会缺项导致失败。
"""

import argparse
import os
import sys
from pathlib import Path

# 复用 skill 目录里的核心模块（单一数据源，避免复制代码）
sys.path.insert(0, str(Path(__file__).resolve().parent
                       / "ucas-schedule-export" / "scripts"))
import ucas_crawler  # noqa: E402


def load_dotenv(path: Path):
    """极简 .env 加载：KEY=VALUE 每行一条，# 注释；已存在的环境变量不覆盖。"""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip("'").strip('"'))


def main():
    load_dotenv(Path(__file__).resolve().parent / ".env")
    ap = argparse.ArgumentParser(
        description="爬取国科大全部学期课表，导出 WakeUp 格式 Excel/CSV"
                    "（登录态取自 UCAS_COOKIE 环境变量或已保存的登录态）")
    ap.add_argument("--out", default=".", help="输出目录（默认当前目录）")
    ap.add_argument("--no-merge", action="store_true", help="教务课程不合并")
    ap.add_argument("--cookie", default="", help="临时使用的 Cookie 字符串")
    args = ap.parse_args()

    if not ucas_crawler.load_cookie_header(args.cookie):
        print("[未登录] 没有可用的登录态。任选其一：")
        print("  1. 设置环境变量 UCAS_COOKIE（可写入本目录 .env 文件，"
              "参考 .env.example）")
        print("  2. 先运行 python3 login_browser.py 完成学习通登录")
        sys.exit(2)

    ucas_crawler.cmd_crawl(argparse.Namespace(
        cookie=args.cookie, out=args.out, no_merge=args.no_merge))


if __name__ == "__main__":
    main()
