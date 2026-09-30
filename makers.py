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
DATE = re.compile(r"(20\d{2})\s*[./年\-]\s*(\d{1,2})(?:\s*[./月\-]\s*(\d{1,2}))?")
JST = datetime.timezone(datetime.timedelta(hours=9))
TODAY = datetime.datetime.now(JST).strftime("%Y%m%d")

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

def _ymd(m):
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
    return f"{y:04d}{mo:02d}{d:02d}" if 1 <= mo <= 12 and 1 <= d <= 31 else ""

def find_date(a, title, depth=6):
    """リンク→親へ遡り、タイトル位置に最も近い日付を YYYYMMDD で返す"""
    cur = a
    for _ in range(depth):
        if cur is None: break
        t = cur.get_text(" ", strip=True)
        ms = list(DATE.finditer(t))
        if ms:
            if len(ms) == 1: return _ymd(ms[0])
            i = t.find(title[:20]) if title else -1
            if i >= 0:
                before = [m for m in ms if m.start() <= i]
                after = [m for m in ms if m.start() > i]
                pick = before[-1] if before else after[0]
                # 前後どちらが近いか（表形式: 日付|タイトル の順が多いので前を優先）
                if before and after and (i - before[-1].end()) > (after[0].start() - i) + 40: pick = after[0]
                return _ymd(pick)
            return ""  # 複数日付でタイトル位置不明 → 断念
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
        if len(title) < 4:  # 「詳細」「PDF」等のリンク文言は親テキストをタイトルに
            title = ctx[:120]
        if len(title) < 6: continue
        if not (KW.search(title) or KW.search(ctx[:200])): continue
        if not STRONG.search(title) and not STRONG.search(ctx[:200]) and not href.lower().endswith(".pdf"): continue
        absu = urljoin(url, href)
        if absu in seen: continue
        seen.add(absu)
        if absu.split("#")[0] == url.split("#")[0]: continue  # 自ページ・カテゴリ切替は除外
        d = find_date(a, title)
        if not d: continue  # 日付が取れないもの（メニュー等）は除外
        if d > TODAY: continue
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
            html = fetch(m["url"])
            items = extract(m["maker"], m["url"], html)
            log(f"{m['maker']}: {len(items)}件")
            for it in items:
                k = it["url"]
                if k not in state:
                    it["first_seen"] = it["date"] if (first_run and it["date"]) else TODAY
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
