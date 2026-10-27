#!/usr/bin/env python3
"""Check events.json before a daily update replaces it.

The 8:40 job should append only new ids, keep every source URL, set status
to 予定 when the year's notice is not final, refresh "updated", then run:

    python3 scripts/validate_events.py
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "events.json"
CATS = {"コンサート", "スポーツ", "展示会", "祭り", "祝日・連休"}
PREFS = {
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
}
DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
NEED = ("id", "date_start", "date_end", "name", "category", "prefecture", "venue", "scale", "source_country", "source", "announced_date", "status")


def months_of(a: str, b: str):
    y, m = int(a[:4]), int(a[5:7])
    end = b[:7]
    out = []
    while True:
        key = f"{y:04d}-{m:02d}"
        out.append(key)
        if key == end:
            break
        m += 1
        if m == 13:
            m = 1
            y += 1
        if y > 2035:
            break
    return out


def main() -> int:
    try:
        doc = json.loads(PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        print(f"events.json を読めません: {e}", file=sys.stderr)
        return 1
    err = []
    items = doc.get("items")
    if not isinstance(items, list):
        err.append("items が配列ではありません")
        items = []
    if not DATE.match(str(doc.get("updated") or "")):
        err.append("updated は YYYY-MM-DD")
    seen = set()
    by_month = Counter()
    by_cat = Counter()
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            err.append(f"#{i} がオブジェクトではありません")
            continue
        miss = [k for k in NEED if k not in it]
        if miss:
            err.append(f"#{i} 不足: {', '.join(miss)}")
        eid = str(it.get("id") or "")
        if not eid:
            err.append(f"#{i} id が空です")
        elif eid in seen:
            err.append(f"id が重複: {eid}")
        seen.add(eid)
        cat = it.get("category")
        if cat not in CATS:
            err.append(f"{eid}: category が不正 ({cat})")
        else:
            by_cat[cat] += 1
        for key in ("date_start", "date_end"):
            if not DATE.match(str(it.get(key) or "")):
                err.append(f"{eid}: {key} が日付ではありません")
        if DATE.match(str(it.get("date_start") or "")) and DATE.match(str(it.get("date_end") or "")):
            if it["date_start"] > it["date_end"]:
                err.append(f"{eid}: date_start が date_end より後")
            else:
                for ym in months_of(it["date_start"], it["date_end"]):
                    if "2026-11" <= ym <= "2027-10":
                        by_month[ym] += 1
        src = str(it.get("source") or "")
        if not src.startswith("http"):
            err.append(f"{eid}: source が URL ではありません")
        if it.get("status") not in ("", "予定"):
            err.append(f"{eid}: status は空か「予定」")
        ad = str(it.get("announced_date") or "")
        if ad and not DATE.match(ad):
            err.append(f"{eid}: announced_date が日付ではありません")
        pref = str(it.get("prefecture") or "")
        if cat == "祝日・連休":
            if pref:
                err.append(f"{eid}: 祝日・連休に都道府県があります")
            if not it.get("source_country"):
                err.append(f"{eid}: source_country が空です")
        elif pref not in PREFS:
            err.append(f"{eid}: prefecture が不正 ({pref})")
        if not str(it.get("name") or "").strip():
            err.append(f"{eid}: name が空です")
    if err:
        print("\n".join(err), file=sys.stderr)
        return 1
    print(f"ok {len(items)}件 updated={doc.get('updated')}")
    print("category")
    for k, n in by_cat.most_common():
        print(f"  {k}\t{n}")
    print("month")
    y, m = 2026, 11
    while (y, m) <= (2027, 10):
        key = f"{y:04d}-{m:02d}"
        print(f"  {key}\t{by_month[key]}")
        m += 1
        if m == 13:
            m = 1
            y += 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
