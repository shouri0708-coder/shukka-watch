# -*- coding: utf-8 -*-
"""メーカー各社の供給情報ページから新着案内（日付・タイトル・リンク）を集め、docs/makers.json を作る。"""
import os, re, json, sys, datetime, time
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

BASE = os.path.dirname(os.path.abspath(__file__))
CONF = os.path.join(BASE, "makers.json")
STATE = os.path.join(BASE, "makers_state.json")
OUT = os.path.join(BASE, "docs", "makers.json")
LOG = os.path.join(BASE, "makers_log.txt")

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36",
      "Accept-Language": "ja,en;q=0.8"}
KW = re.compile(r"供給|出荷|限定|停止|中止|再開|休止|回収|欠品|集約|辞退|終了|一時|安定|案内|お知らせ|変更|発売")
STRONG = re.compile(r"供給|出荷|限定|停止|中止|再開|休止|回収|欠品|集約|辞退|終了")
SUPPLY = re.compile(r"供給|出荷|限定|停止|中止|再開|休止|欠品|集約|辞退|販売終了|製造終了|回収")
NAVQ = re.compile(r"[?&](cat|cate|category|news-type|type|kind|tag|page|p)=", re.I)
NOISE = re.compile(r"^(新着情報|お知らせ|安全性|添付文書|電子添文|包装変更|販売中止・|流通情報|供給・中止|供給関連|発売情報|その他|重要なお知らせ)")
DATE_TOK = re.compile(r"20\d{2}\s*[./年\-]\s*\d{1,2}(\s*[./月\-]\s*\d{1,2})?\s*日?")
LABEL = re.compile(r"^(供給に関するお知らせ|供給関連|発売情報|流通情報・回収情報|販売中止・経過措置|安全性情報|電子添文改訂|重要なお知らせ|お知らせ|情報)\s*")
LABEL2 = re.compile(r"^(販売中止|供給関連情報|供給関連|新発売|包装変更|中止・経過措置|重要なお知らせ|重要|供給)\s+(?=\S)")
TAIL = re.compile(r"(お知らせ文書)?を掲載(いた)?しました。?$")
DATE = re.compile(r"(20\d{2})\s*[./年\-]\s*(\d{1,2})(?:\s*[./月\-]\s*(\d{1,2}))?")
JST = datetime.timezone(datetime.timedelta(hours=9))
TODAY = datetime.datetime.now(JST).strftime("%Y%m%d")

def clean_title(t):
    t = DATE_TOK.sub(" ", t)
    t = re.sub(r"\b(PDF|NEW|New|new)\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip(" 　:：・|")
    for _ in range(3): t = LABEL2.sub("", LABEL.sub("", t)).strip(" 　:：・|")
    t = TAIL.sub("", t).strip(" 　:：・|")
    t = re.sub(r"[（(]\s*[^（()）]*向け[^（()）]*[)）]\s*$", "", t).strip()  # （医療関係者様向け | 特約店様向け）等
    return t

def log(*a):
    s = " ".join(str(x) for x in a); print(s)
    with open(LOG, "a", encoding="utf-8") as f: f.write(s + "\n")

def fetch(url):
    r = requests.get(url, headers=UA, timeout=30)
    r.raise_for_status()
    b = r.content
    for enc in (r.encoding if r.encoding and r.encoding.lower() != "iso-8859-1" else None, "utf-8", "cp932", "euc_jp"):
        if not enc: continue
        try: return b.decode(enc)
        except Exception: pass
    return b.decode("utf-8", "replace")

GATE = r"はい|医療関係者(です|の方|用)|医療従事者|同意(する|して)|閲覧する|入る|進む|確認しました|OK|承諾|Yes"
def fetch_browser(url, wait_ms=3000, gate=None):
    """JS描画・確認ゲートのあるページを Playwright で取得"""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(user_agent=UA["User-Agent"], locale="ja-JP", viewport={"width": 1280, "height": 2000})
        pg = ctx.new_page()
        pg.goto(url, wait_until="domcontentloaded", timeout=60000)
        pg.wait_for_timeout(wait_ms)
        js = """(re)=>{const r=new RegExp(re);const els=[...document.querySelectorAll('a,button,input[type=button],input[type=submit],label,div[role=button],span[role=button]')];
                 for(const e of els){const t=(e.innerText||e.value||'').trim();if(t&&t.length<=24&&r.test(t)){e.click();return t;}}return '';}"""
        for _ in range(3):
            try: hit = pg.evaluate(js, gate or GATE)
            except Exception: hit = ""
            if not hit: break
            log(f"  gate click: {hit}")
            pg.wait_for_timeout(2500)
        if pg.url.split("#")[0] != url.split("#")[0]:
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(wait_ms)
            except Exception: pass
        # 「もっと見る」系を数回押して一覧を伸ばす
        for _ in range(3):
            try: more = pg.evaluate(js, r"^(もっと見る|さらに表示|More|次へ|すべて表示)$")
            except Exception: more = ""
            if not more: break
            pg.wait_for_timeout(1500)
        html = pg.content(); b.close()
    return html

def _ymd(m):
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
    return f"{y:04d}{mo:02d}{d:02d}" if 1 <= mo <= 12 and 1 <= d <= 31 else ""

def find_date(a, title, depth=6, maxdist=160):
    """リンク→親へ遡り、タイトル位置から maxdist 文字以内で最も近い日付を YYYYMMDD で返す"""
    cur = a
    for _ in range(depth):
        if cur is None: break
        t = cur.get_text(" ", strip=True)
        ms = list(DATE.finditer(t))
        if ms:
            i = t.find(title[:20]) if title else -1
            if i < 0: i = 0
            best, bd = None, 10**9
            for m in ms:
                dist = (i - m.end()) if m.end() <= i else (m.start() - (i + len(title[:20])))
                if dist < 0: dist = 0
                if m.end() > i: dist += 40  # 後ろの日付より前の日付を優先
                if dist < bd: best, bd = m, dist
            return _ymd(best) if bd <= maxdist else ""
        cur = cur.parent
    return ""

def extract(maker, url, html):
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "nav", "header", "footer"]): t.decompose()
    seen, out = set(), []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith(("#", "javascript:", "mailto:")): continue
        title = a.get_text(" ", strip=True)
        ctx = a.parent.get_text(" ", strip=True) if a.parent else title
        if len(title) < 4 or (len(title) <= 14 and not SUPPLY.search(title) and SUPPLY.search(ctx)):
            title = ctx[:140]  # 「詳細」「PDF」「医療関係者様向け」等のリンク文言は親テキストをタイトルに
        if len(title) < 6: continue
        if not (KW.search(title) or KW.search(ctx[:200])): continue
        if not STRONG.search(title) and not STRONG.search(ctx[:200]) and not href.lower().endswith(".pdf"): continue
        absu = urljoin(url, href)
        if absu in seen: continue
        seen.add(absu)
        if absu.split("#")[0] == url.split("#")[0]: continue  # 自ページ・カテゴリ切替は除外
        if NAVQ.search(absu) and not re.search(r"[?&](id|if_id|news_id|nid)=", absu): continue
        if not SUPPLY.search(title): continue
        title = clean_title(title)
        if len(title) < 8 or NOISE.match(title): continue
        d = find_date(a, title)
        if not d: continue  # 日付が取れないもの（メニュー等）は除外
        if d > TODAY: continue
        dk = (maker, d, title[:60])
        if dk in seen: continue  # 同日・同タイトル（医療関係者向け/特約店向け 等）は1件に
        seen.add(dk)
        out.append({"maker": maker, "title": title[:140], "url": absu, "date": d, "src": url})
    return out

def main():
    open(LOG, "w").close()
    conf = json.load(open(CONF, encoding="utf-8"))
    state = json.load(open(STATE, encoding="utf-8")) if os.path.exists(STATE) else {}
    first_run = not state
    new_cnt = 0
    for m in conf:
        if not m.get("enabled", True): continue
        try:
            html = fetch_browser(m["url"], gate=m.get("gate")) if m.get("mode") == "browser" else fetch(m["url"])
            if os.environ.get("MAKERS_DEBUG"):
                os.makedirs(os.path.join(BASE, "debug"), exist_ok=True)
                open(os.path.join(BASE, "debug", f"{m['maker']}.html"), "w", encoding="utf-8").write(html)
            items = extract(m["maker"], m["url"], html)
            log(f"{m['maker']}: {len(items)}件")
            for it in items:
                k = it["url"]
                if k not in state:
                    it["first_seen"] = it["date"] or TODAY
                    state[k] = it; new_cnt += 1
                else:
                    state[k]["title"] = it["title"]; state[k]["date"] = it["date"] or state[k].get("date", "")
            time.sleep(1)
        except Exception as e:
            log(f"{m['maker']}: ERROR {type(e).__name__}: {e}")
    # 出力: 直近120日（日付 or 初見日）
    lim = (datetime.datetime.now(JST) - datetime.timedelta(days=120)).strftime("%Y%m%d")
    items = [v for v in state.values() if max(v.get("date", ""), v.get("first_seen", "")) >= lim]
    items.sort(key=lambda v: max(v.get("date", ""), v.get("first_seen", "")), reverse=True)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump({"updated": TODAY, "items": items}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False)
    json.dump(state, open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    log(f"新規 {new_cnt}件 / 出力 {len(items)}件")

if __name__ == "__main__":
    main()
