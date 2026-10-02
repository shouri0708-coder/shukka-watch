# -*- coding: utf-8 -*-
"""東和薬品の「全製品供給状況一覧」Excel を職種選択ゲートを通って取得する。"""
import os, re, sys
from playwright.sync_api import sync_playwright
URL = "https://med.towayakuhin.co.jp/medical/product/information_supply.php"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"

def fetch(dest):
    with sync_playwright() as p:
        b = p.chromium.launch(); ctx = b.new_context(user_agent=UA, locale="ja-JP"); pg = ctx.new_page()
        pg.goto(URL, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(2000)
        if "job_selector" in pg.content():
            pg.click("text=薬剤師（薬局）"); pg.wait_for_timeout(2500)
            if "information_supply" not in pg.url: pg.goto(URL, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(2000)
        html = pg.content()
        m = re.search(r'href="([^"]*fileloader\.php\?id=\d+&(?:amp;)?t=\d+)"[^>]*>[^<]*(?:<[^>]+>)*\s*(20\d{2}年\d{1,2}月\d{1,2}日)更新【Excel】', html)
        if not m:
            # 緩い探索
            links = re.findall(r'(20\d{2}年\d{1,2}月\d{1,2}日)更新【Excel】', html); hrefs = re.findall(r'href="([^"]*fileloader\.php\?id=\d+[^"]*)"', html)
            print("Excelリンク特定失敗", links[:2], hrefs[:4]); b.close(); return None
        href, upd = m.group(1).replace("&amp;", "&"), m.group(2)
        if href.startswith("/"): href = "https://med.towayakuhin.co.jp" + href
        r = ctx.request.get(href, timeout=120000)
        data = r.body(); b.close()
        open(dest, "wb").write(data)
        print("saved", dest, len(data), "bytes; 更新:", upd, href)
        return upd

if __name__ == "__main__":
    fetch(sys.argv[1] if len(sys.argv) > 1 else "towa.xlsx")
