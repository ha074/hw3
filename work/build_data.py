# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1",
# ]
# ///
"""
Build docs/data.js for the BI page (docs/index.html) from 中低1.0.xlsx.

The output is a plain <script>-loadable JS file (window.BI_DATA = {...}), so
index.html works when opened directly from disk (file://) without a server.

Record fields (one row per finest category actually present in the report):
  year          int     e.g. 2026
  quarter       str     "Q1".."Q4"
  income_type   str     "middle_low_income" | "low_income"
  county        str     county (zh) the township is listed under in the report
  region       str     township name (zh), e.g. "花蓮市"
  region_en     str     township name (en), as printed in the report
  measure       str     "households" | "persons"
  sex           str     "male" | "female"
  age_group     str|null  "under_12" | "age_12_17" | "age_18_64" | "age_65_over";
                         null for households (the report has no age split for 戶數)
  value         int

Excluded (aggregates -> would double count): the 總計 row, the 花蓮縣 county row,
all 計 / 總計 subtotal columns. Quarters whose cells are all empty are skipped.
The workbook is opened read-only and never written.

Usage: uv run work/build_data.py [source.xlsx] [output.js]
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SRC = ROOT / "中低1.0.xlsx"
DEFAULT_OUT = ROOT / "docs" / "data.js"

INCOME_TYPES = [("Middle-Low-Income", "middle_low_income", "中低收入戶"),
                ("Low-Income", "low_income", "低收入戶")]
SEX = {"計": "subtotal", "總計": "subtotal", "男": "male", "女": "female"}
AGE = {"計": "subtotal", "未滿12歲": "under_12", "12-17歲": "age_12_17",
       "18-64歲": "age_18_64", "65歲以上": "age_65_over"}
LABELS = {
    "income_type": {"middle_low_income": "中低收入戶", "low_income": "低收入戶"},
    "measure": {"households": "戶數(戶)", "persons": "人數(人)"},
    "sex": {"male": "男", "female": "女"},
    "age_group": {"under_12": "未滿12歲", "age_12_17": "12-17歲",
                  "age_18_64": "18-64歲", "age_65_over": "65歲以上"},
}


def load_rows(ws):
    # read_only worksheets don't expose merged ranges, so the workbook is opened
    # normally; nothing is ever saved.
    merged = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged[(r, c)] = (rng.min_row, rng.min_col)

    def hv(r, c):  # header value with merged-cell fill
        r0, c0 = merged.get((r, c), (r, c))
        v = ws.cell(r0, c0).value
        return None if v is None else str(v).strip()

    header_rows = [r for r in range(1, ws.max_row + 1)
                   if str(ws.cell(r, 1).value or "").startswith("區域別")]
    records = []
    for i, h in enumerate(header_rows):
        # report title (search upwards)
        title = next(str(ws.cell(r, 1).value) for r in range(h - 1, 0, -1)
                     if "Number and Population" in str(ws.cell(r, 1).value or ""))
        income = next(code for key, code, _ in INCOME_TYPES if key in title
                      and not (key == "Low-Income" and "Middle-Low-Income" in title))
        end = header_rows[i + 1] - 1 if i + 1 < len(header_rows) else ws.max_row

        # column map from the 4 header levels
        cols = []
        for c in range(3, ws.max_column + 1):
            q = hv(h, c)
            if not q:
                continue
            mq, my = re.search(r"第\s*(\d)\s*季", q), re.search(r"(\d{4})", q)
            meas_lbl = hv(h + 1, c)
            measure = "households" if meas_lbl.startswith("戶數") else \
                      "persons" if meas_lbl.startswith("人數") else None
            if not (mq and my and measure):
                raise ValueError(f"Unrecognised header at row {h}, col {c}: {q!r} / {meas_lbl!r}")
            sex = SEX[hv(h + 2, c).split("\n")[0].split()[0]]
            age = AGE[hv(h + 3, c).split("\n")[0]] if measure == "persons" else None
            cols.append((c, int(my.group(1)), f"Q{mq.group(1)}", measure, sex, age))

        # data rows: township rows only (skip 總計 / county aggregates).
        # The county a township belongs to is the county row printed above it
        # in the same block (report hierarchy), not a guess.
        data, county = [], None
        for r in range(h + 5, end + 1):
            zh, en = ws.cell(r, 1).value, ws.cell(r, 2).value
            if zh is None and en is None:
                break
            zh, en = str(zh).strip(), str(en).strip()
            if en == "Grand-Total":
                continue
            if "County" in en:
                county = zh
                continue
            if county is None:
                raise ValueError(f"Township row {r} ({zh}) has no county row above it")
            data.append((r, zh, en, county))

        # skip quarters that are entirely empty in the source
        filled = {q for c, _, q, *_ in cols
                  if any(ws.cell(r, c).value is not None for r, *_ in data)}
        for r, zh, en, county in data:
            for c, year, q, measure, sex, age in cols:
                if q not in filled or sex == "subtotal" or age == "subtotal":
                    continue
                v = ws.cell(r, c).value
                if v is None or not isinstance(v, (int, float)) or v != int(v):
                    raise ValueError(f"Bad value {v!r} at row {r}, col {c}")
                records.append({"year": year, "quarter": q, "income_type": income,
                                "county": county, "region": zh, "region_en": en, "measure": measure,
                                "sex": sex, "age_group": age, "value": int(v)})
    return records


def aggregate(records):
    """Sum records sharing the same category key (keeps first-seen order)."""
    key_fields = ["year", "quarter", "income_type", "county", "region", "region_en",
                  "measure", "sex", "age_group"]
    acc = OrderedDict()
    for rec in records:
        k = tuple(rec[f] for f in key_fields)
        acc[k] = acc.get(k, 0) + rec["value"]
    return [dict(zip(key_fields, k), value=v) for k, v in acc.items()]


def git_repo_url(root: Path):
    """Return (https_url, None) from the project's Git remote, or (None, reason).

    Only a real github.com remote is accepted; nothing is guessed.
    """
    try:
        res = subprocess.run(["git", "-C", str(root), "remote"], capture_output=True,
                             text=True, timeout=10)
    except (OSError, subprocess.SubprocessError):
        return None, "找不到 git 指令"
    if res.returncode != 0:
        return None, "專案資料夾不是 Git repository"
    remotes = res.stdout.split()
    if not remotes:
        return None, "Git repository 沒有設定 remote"
    name = "origin" if "origin" in remotes else remotes[0]
    url = subprocess.run(["git", "-C", str(root), "remote", "get-url", name],
                         capture_output=True, text=True, timeout=10).stdout.strip()
    m = (re.fullmatch(r"git@github\.com:([\w.-]+)/([\w.-]+?)(?:\.git)?/?", url)
         or re.fullmatch(r"(?:ssh://git@|https?://(?:[^@/]+@)?)github\.com[:/]([\w.-]+)/([\w.-]+?)(?:\.git)?/?", url))
    if not m:
        return None, f"remote「{name}」不是有效的 GitHub 網址"
    return f"https://github.com/{m.group(1)}/{m.group(2)}", None


def main(src: Path, out: Path):
    wb = openpyxl.load_workbook(src, data_only=True)
    raw = [rec for ws in wb.worksheets for rec in load_rows(ws)]
    wb.close()
    records = aggregate(raw)

    def uniq(f):
        return list(dict.fromkeys(r[f] for r in records))

    repo_url, repo_reason = git_repo_url(ROOT)
    payload = {
        "meta": {
            "source": src.name,
            "publisher": "衛生福利部統計處",
            "repo_url": repo_url,
            "repo_missing_reason": repo_reason,
            "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "record_count": len(records),
            "fields": ["year", "quarter", "income_type", "county", "region", "region_en",
                       "measure", "sex", "age_group", "value"],
            "notes": [
                "age_group is null for households: the report gives household counts by sex only.",
                "Excluded aggregates: 總計 row, 花蓮縣 county row, 計/總計 subtotal columns.",
                "Quarters with headers but no data in the source are not included.",
            ],
            "dimensions": {
                "year": uniq("year"), "quarter": uniq("quarter"),
                "income_type": uniq("income_type"), "county": uniq("county"),
                "region": uniq("region"),
                "measure": uniq("measure"), "sex": uniq("sex"),
                "age_group": [a for a in uniq("age_group") if a is not None],
            },
            "region_en": {r["region"]: r["region_en"] for r in records},
            "labels": LABELS,
        },
        "records": records,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    out.write_text("// Generated by work/build_data.py — do not edit by hand.\n"
                   f"window.BI_DATA = {body};\n", encoding="utf-8")
    print(f"raw cells: {len(raw)}, records after aggregation: {len(records)}")
    print(f"wrote {out} ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SRC
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_OUT
    main(src, out)
