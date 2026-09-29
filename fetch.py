# -*- coding: utf-8 -*-
"""厚労省「医療用医薬品供給状況」Excelを取得して data/ に保存する。"""
import re, sys, os, urllib.request, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) shukka-watch/1.0"}

def get(url, binary=False):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=120) as r:
        b = r.read()
    return b if binary else b.decode("utf-8", "replace")

def main():
    os.makedirs(DATA, exist_ok=True)
    html = get(PAGE)
    links = re.findall(r'href="([^"]*?/content/\d+/(\d{6})iyakuhinkyoukyu[^"]*?\.xlsx)"', html)
    if not links:
        print("Excelリンクが見つかりません。ページ構成が変わった可能性があります。")
        sys.exit(1)
    # 日付6桁(YYMMDD)が最大のものを採用
    href, ymd = max(links, key=lambda t: t[1])
    if href.startswith("/"):
        href = "https://www.mhlw.go.jp" + href
    date = "20" + ymd  # YYYYMMDD
    dest = os.path.join(DATA, f"{date}.xlsx")
    if os.path.exists(dest) and os.path.getsize(dest) > 0:
        print(f"取得済み: {dest}")
    else:
        print(f"ダウンロード: {href}")
        b = get(href, binary=True)
        with open(dest, "wb") as f:
            f.write(b)
        print(f"保存: {dest} ({len(b)//1024} KB)")
    with open(os.path.join(DATA, "latest.txt"), "w", encoding="utf-8") as f:
        f.write(date)
    print("公表日:", date)

if __name__ == "__main__":
    main()
