# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1",
# ]
# ///
"""
Verify docs/data.js against 中低1.0.xlsx.

data.js is executed with Node (as a browser would run it) to obtain window.BI_DATA.
The workbook is re-read independently using the report's fixed layout (each quarter
= 18 columns: 戶數 計/男/女, 人數 總計x5 / 男x5 / 女x5); header labels are asserted.

Checks: every value equals its source cell, no missing / extra / duplicate records,
detail sums reproduce the report's 計/總計 columns and the 花蓮縣 county row,
zero values kept, grand totals per quarter/income type/measure match.

Usage: uv run work/verify_data.py [source.xlsx] [data.js]   (needs `node` on PATH)
"""
from __future__ import annotations

import json
import subprocess
import sys
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
AGES = ["under_12", "age_12_17", "age_18_64", "age_65_over"]
AGE_ZH = ["未滿12歲", "12-17歲", "18-64歲", "65歲以上"]
QSTART = {"Q1": 3, "Q2": 21, "Q3": 39, "Q4": 57}
BLOCKS = [("middle_low_income", 3, 8, 22), ("low_income", 26, 31, 44)]

fails, lines = 0, []


def check(name, ok, detail=""):
    global fails
    fails += not ok
    lines.append(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))


def load_js(path):
    js = ("const vm=require('vm');const fs=require('fs');const ctx={window:{}};"
          f"vm.runInNewContext(fs.readFileSync({json.dumps(str(path))},'utf8'),ctx);"
          "process.stdout.write(JSON.stringify(ctx.window.BI_DATA));")
    return json.loads(subprocess.run(["node", "-e", js], check=True,
                                     capture_output=True, text=True).stdout)


def main(src: Path, js: Path):
    data = load_js(js)
    check("data.js defines window.BI_DATA", isinstance(data, dict) and "records" in data)
    recs = data["records"]
    fields = ["year", "quarter", "income_type", "county", "region", "region_en", "measure",
              "sex", "age_group", "value"]
    check("record fields", all(list(r) == fields for r in recs))
    check("meta.record_count matches", data["meta"]["record_count"] == len(recs))
    check("values are non-negative integers",
          all(isinstance(r["value"], int) and r["value"] >= 0 for r in recs))
    check("no missing values (age_group null only for households)",
          all(all(r[f] not in (None, "") for f in fields if f != "age_group") for r in recs)
          and all((r["age_group"] is None) == (r["measure"] == "households") for r in recs))
    key = lambda r: (r["year"], r["quarter"], r["income_type"], r["region"],
                     r["measure"], r["sex"], r["age_group"])
    dups = [k for k, n in Counter(map(key, recs)).items() if n > 1]
    check("no duplicate category keys", not dups, f"{len(dups)} duplicates")
    check("no aggregate regions (總計/花蓮縣)",
          not {r["region"] for r in recs} & {"總計", "花蓮縣"})
    idx = {key(r): r["value"] for r in recs}

    ws = openpyxl.load_workbook(src, data_only=True).worksheets[0]
    v = lambda r, c: ws.cell(r, c).value
    expected, mism, sub_mism, county_mism, zeros, totals = set(), [], [], [], 0, []
    for income, h, r0, r1 in BLOCKS:
        rows = {(str(v(r, 1)).strip(), str(v(r, 2)).strip()): r for r in range(r0, r1 + 1)}
        county = next(r for (zh, en), r in rows.items() if "County" in en)
        county_zh = str(v(county, 1)).strip()
        bad_cty = [r for r in recs if r["income_type"] == income and r["county"] != county_zh]
        check(f"[{income}] county field = {county_zh} (row above the townships)", not bad_cty,
              f"{len(bad_cty)} records differ")
        towns = {zh: r for (zh, en), r in rows.items() if en != "Grand-Total" and "County" not in en}
        for q, c0 in QSTART.items():
            if v(h, c0) is None:
                continue
            if all(v(r, c0 + k) is None for r in range(r0, r1 + 1) for k in range(18)):
                lines.append(f"[INFO] {income} {q}: header present, no data in source (not expected in output)")
                continue
            # layout guard
            assert f"第{q[1]}季" in v(h, c0).replace(" ", "")
            assert [str(v(h + 2, c0 + k)).split("\n")[0] for k in range(3)] == ["計", "男", "女"]
            for b in (3, 8, 13):
                assert [v(h + 3, c0 + b + k) for k in range(5)] == ["計"] + AGE_ZH
            year = int(str(v(h, c0)).split(",")[-1].split()[0])
            for zh, r in towns.items():
                cells = [("households", "male", None, 1), ("households", "female", None, 2)]
                cells += [("persons", sx, a, base + 1 + i)
                          for sx, base in (("male", 8), ("female", 13)) for i, a in enumerate(AGES)]
                got = {}
                for m, sx, a, off in cells:
                    k = (year, q, income, zh, m, sx, a)
                    expected.add(k)
                    sv = v(r, c0 + off)
                    zeros += sv == 0
                    got[(m, sx, a)] = idx.get(k)
                    if idx.get(k) != sv:
                        mism.append(f"{income} {q} {zh} {m}/{sx}/{a}: js={idx.get(k)} xlsx={sv}")
                if None in got.values():
                    continue
                hh = got[("households", "male", None)] + got[("households", "female", None)]
                if hh != v(r, c0):
                    sub_mism.append(f"{income} {q} {zh} 戶數計")
                for sx, base in (("male", 8), ("female", 13)):
                    if sum(got[("persons", sx, a)] for a in AGES) != v(r, c0 + base):
                        sub_mism.append(f"{income} {q} {zh} 人數{sx}計")
                for i, a in enumerate(AGES):
                    if got[("persons", "male", a)] + got[("persons", "female", a)] != v(r, c0 + 4 + i):
                        sub_mism.append(f"{income} {q} {zh} 總計/{a}")
                if sum(x for (m, *_), x in got.items() if m == "persons") != v(r, c0 + 3):
                    sub_mism.append(f"{income} {q} {zh} 人數總計")
            for k in range(18):
                if sum(v(r, c0 + k) for r in towns.values()) != v(county, c0 + k):
                    county_mism.append(f"{income} {q} col+{k}")
            # report-level totals: data.js vs 花蓮縣 row
            for m, col in (("households", c0), ("persons", c0 + 3)):
                js_total = sum(r["value"] for r in recs if r["income_type"] == income
                               and r["quarter"] == q and r["measure"] == m)
                totals.append((income, q, m, js_total, v(county, col)))

    check("every value equals its source cell", not mism, "; ".join(mism[:5]))
    missing, extra = expected - set(idx), set(idx) - expected
    check("no records missing", not missing, f"{len(missing)} missing")
    check("no extra records", not extra, f"{len(extra)} extra")
    check("record count", len(recs) == len(expected), f"data.js={len(recs)} expected={len(expected)}")
    check("detail sums = report 計/總計 columns", not sub_mism, "; ".join(sub_mism[:5]))
    check("township sums = 花蓮縣 county row", not county_mism, "; ".join(county_mism[:5]))
    zjs = sum(r["value"] == 0 for r in recs)
    check("zero values kept", zjs == zeros, f"xlsx={zeros} data.js={zjs}")
    for income, q, m, a, b in totals:
        check(f"total {income} {q} {m}", a == b, f"data.js={a:,} 花蓮縣={b:,}")

    print("\n".join(lines))
    print(f"\nRESULT: {'PASSED' if fails == 0 else f'FAILED ({fails})'}")
    return 0 if fails == 0 else 1


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "中低1.0.xlsx"
    js = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "docs" / "data.js"
    sys.exit(main(src, js))
