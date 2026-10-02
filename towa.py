# -*- coding: utf-8 -*-
"""東和薬品: 「全製品供給状況一覧」Excel の『更新履歴』シートを品目ごとの新着に変換する。"""
import re, io, datetime
import openpyxl

URL = "https://med.towayakuhin.co.jp/medical/product/information_supply.php"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
PROD = re.compile(r"「(トーワ|TW|ＴＷ|東和)」|錠|カプセル|散|顆粒|細粒|シロップ|液|注|軟膏|クリーム|テープ|パップ|ゲル|点眼|坐剤|吸入|貼付|ドライシロップ|配合")

def fetch_excel(log=print):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(); ctx = b.new_context(user_agent=UA, locale="ja-JP"); pg = ctx.new_page()
        pg.goto(URL, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(2000)
        if "job_selector" in pg.content():
            pg.click("text=薬剤師（薬局）"); pg.wait_for_timeout(2500)
            if "information_supply" not in pg.url:
                pg.goto(URL, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(2000)
        html = pg.content()
        m = re.search(r'href="([^"]*fileloader\.php\?id=\d+&(?:amp;)?t=\d+)"[^>]*>[^<]*(?:<[^>]+>)*\s*(20\d{2}年\d{1,2}月\d{1,2}日)更新【Excel】', html)
        if not m:
            b.close(); raise RuntimeError("東和: Excelリンクが見つかりません")
        href = m.group(1).replace("&amp;", "&")
        if href.startswith("/"): href = "https://med.towayakuhin.co.jp" + href
        data = ctx.request.get(href, timeout=120000).body(); b.close()
    log(f"  東和 Excel {len(data)//1024}KB ({m.group(2)}更新)")
    return data

def parse_history(data, maker="東和薬品"):
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = None
    for w in wb.worksheets:
        if "更新履歴" in w.title: ws = w; break
    if ws is None: return []
    items, date, cat = [], "", ""
    for row in ws.iter_rows(min_row=3, values_only=True):
        cells = ["" if c is None else str(c).strip() for c in row]
        c0 = cells[0] if len(cells) > 0 else ""
        c1 = cells[1] if len(cells) > 1 else ""
        if c0:
            m = re.match(r"(\d{4})-(\d{2})-(\d{2})", c0) or re.match(r"(\d{4})[/.](\d{1,2})[/.](\d{1,2})", c0)
            if m: date = "%04d%02d%02d" % tuple(int(x) for x in m.groups())
        if c1:
            c1n = re.sub(r"\s+", "", c1)
            cat = (cat + c1n) if c1n.startswith("（") and cat else c1n
        if not date: continue
        detail = ""
        for c in cells[3:4]:
            if c.startswith("：") or "から「" in c: detail = c.lstrip("：:")
        for c in cells[2:]:
            if not c or c.startswith("以下") or c.startswith("：") or "から「" in c: continue
            name = re.sub(r"\s+", " ", c).strip()
            if not PROD.search(name) or len(name) < 4: continue
            if re.fullmatch(r"（?リンク更新）?|その他の変更（リンク更新）", cat): continue
            kind = cat
            mm = re.search(r"（(.+?)）", cat)
            if cat.startswith("その他の変更") and mm: kind = mm.group(1)
            if "リンク" in kind and "から「" not in detail: continue
            mv = re.search(r"「(.+?)」から「(.+?)」に変更", detail) if detail else None
            if mv:
                def ship(x):
                    ps = [q.strip() for q in re.split(r"[、,]", x)]
                    sp = [q for q in ps if re.match(r"[①-⑤]", q)]
                    return sp[0] if sp else ps[-1]
                title = f"{name} {kind}（{ship(mv.group(1))} → {ship(mv.group(2))}）"
            else:
                title = f"{name} {kind}"
            items.append({"maker": maker, "title": title[:160], "url": URL + f"#{date}-{name}", "date": date, "src": URL, "first_seen": date})
    return items

if __name__ == "__main__":
    import sys
    data = open(sys.argv[1], "rb").read() if len(sys.argv) > 1 else fetch_excel()
    for it in parse_history(data)[:40]: print(it["date"], it["title"])
