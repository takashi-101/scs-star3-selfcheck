"""SCS評価制度「★3・★4要求事項及び評価基準」(Excel) を JSON に変換する。

標準ライブラリだけで動く（閉域端末でも実行できるように openpyxl を使わない）。

使い方:
    python scripts/convert_requirements.py data/source/scs_3.xlsx \
        --source-url https://www.meti.go.jp/policy/netsecurity/docs/scs/scs_3.xlsx \
        --retrieved 2026-10-27 --out data/requirements.json

要求事項・評価基準の文言は Excel のセルの値をそのまま使い、書き換えない。
（ふりがな情報 <rPh> は本文ではないので取り除く。）
"""
import argparse
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# 見出し行の文言 → 内部の項目名。列の位置ではなく見出しの文言で列を特定する。
HEADER_MAP = {
    "大分類No.": "domain_no",
    "大分類": "domain",
    "中分類No.": "category_no",
    "中分類": "category",
    "要求事項No.": "req_no",
    "★3": "req_star3",
    "★4": "req_star4",
    "要求事項名": "req_title",
    "要求事項": "req_text",
    "★3/★4": "crit_level",
    "評価基準No.": "crit_no",
    "評価基準": "crit_text",
    "NIST CSFにおける機能": "nist_csf",
}
REQUIRED_COLUMNS = ["domain_no", "domain", "category_no", "category", "req_no",
                    "req_star3", "req_title", "req_text", "crit_level", "crit_no",
                    "crit_text"]

REQ_NO_RE = re.compile(r"^\d+-\d+-\d+$")
CRIT_NO_RE = re.compile(r"^\d+-\d+-\d+-\d+$")


class ConversionError(Exception):
    pass


def _text_of(si):
    """<si> / <is> 要素から本文だけを取り出す（<rPh> のふりがなは除く）。"""
    parts = [t.text or "" for t in si.findall("m:t", NS)]
    parts += [t.text or "" for t in si.findall("m:r/m:t", NS)]
    return "".join(parts)


def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return _col_index(m.group(1)), int(m.group(2))


def read_sheet(path, sheet_name=None):
    """先頭（または指定名）のシートを {(行, 列): 文字列} で返す。結合セルは左上の値で埋める。"""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [_text_of(si) for si in root.findall("m:si", NS)]

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = wb.findall("m:sheets/m:sheet", NS)
        if not sheets:
            raise ConversionError("ブックにシートがありません")
        sheet = sheets[0]
        if sheet_name is not None:
            matches = [s for s in sheets if s.get("name") == sheet_name]
            if not matches:
                raise ConversionError(f"シート {sheet_name!r} がありません")
            sheet = matches[0]
        rid = sheet.get("{%s}id" % NS["r"])
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels.findall("{%s}Relationship" % REL_NS)
                      if r.get("Id") == rid)
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        ws = ET.fromstring(z.read(target))

    cells = {}
    for c in ws.findall("m:sheetData/m:row/m:c", NS):
        t = c.get("t")
        v = c.find("m:v", NS)
        if t == "inlineStr":
            is_ = c.find("m:is", NS)
            val = _text_of(is_) if is_ is not None else ""
        elif v is None:
            continue
        elif t == "s":
            val = shared[int(v.text)]
        else:
            val = v.text or ""
        cells[_split_ref(c.get("r"))[::-1]] = val

    for mc in ws.findall("m:mergeCells/m:mergeCell", NS):
        a, b = mc.get("ref").split(":")
        c1, r1 = _split_ref(a)
        c2, r2 = _split_ref(b)
        top_left = cells.get((r1, c1))
        if top_left is None:
            continue
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cells[(r, c)] = top_left
    return cells, sheet.get("name")


def _norm(s):
    return (s or "").strip()


def find_header(cells):
    """「評価基準No.」を含む行を見出し行とみなし、{項目名: 列番号} を返す。"""
    rows = sorted({r for r, _ in cells})
    for r in rows:
        row = {c: _norm(v) for (rr, c), v in cells.items() if rr == r}
        if "評価基準No." in row.values() and "要求事項No." in row.values():
            colmap = {}
            for c, label in row.items():
                if label in HEADER_MAP:
                    colmap[HEADER_MAP[label]] = c
            missing = [k for k in REQUIRED_COLUMNS if k not in colmap]
            if missing:
                raise ConversionError(f"見出しに必要な列がありません: {missing}")
            return r, colmap
    raise ConversionError("見出し行（要求事項No.／評価基準No.）が見つかりません")


def convert(cells):
    header_row, col = find_header(cells)
    last_row = max(r for r, _ in cells)

    def get(r, key):
        c = col.get(key)
        return _norm(cells.get((r, c))) if c else ""

    domains, requirements, criteria = {}, {}, []
    for r in range(header_row + 1, last_row + 1):
        crit_no = get(r, "crit_no")
        if not crit_no:
            continue
        if not CRIT_NO_RE.match(crit_no):
            raise ConversionError(f"{r}行目: 評価基準No.の形式が不正です: {crit_no!r}")
        req_no = get(r, "req_no")
        if not REQ_NO_RE.match(req_no):
            raise ConversionError(f"{r}行目: 要求事項No.の形式が不正です: {req_no!r}")

        dom_no = get(r, "domain_no")
        domains.setdefault(dom_no, {"no": dom_no, "name": get(r, "domain")})

        req = {
            "no": req_no,
            "domain_no": dom_no,
            "category_no": get(r, "category_no"),
            "category": get(r, "category"),
            "title": get(r, "req_title"),
            "text": get(r, "req_text"),
            "star3": bool(get(r, "req_star3")),
            "star4": bool(get(r, "req_star4")),
            "nist_csf": get(r, "nist_csf"),
        }
        if req_no in requirements:
            prev = requirements[req_no]
            if (prev["title"], prev["text"]) != (req["title"], req["text"]):
                raise ConversionError(f"{r}行目: 要求事項No.{req_no} の内容が他の行と食い違っています")
            # 元の表では★3/★4の○が評価基準の行ごとに付いている（★4だけの評価基準の行は★3欄が空）。
            # 要求事項としては、いずれかの行に○があればその段階の対象とする。
            prev["star3"] = prev["star3"] or req["star3"]
            prev["star4"] = prev["star4"] or req["star4"]
        else:
            requirements[req_no] = req

        level = get(r, "crit_level")
        if level not in ("★3", "★4"):
            raise ConversionError(f"{r}行目: ★3/★4 の値が不正です: {level!r}")
        criteria.append({
            "no": crit_no,
            "req_no": req_no,
            "level": level,
            "text": get(r, "crit_text"),
            "source_row": r,
        })
    return list(domains.values()), list(requirements.values()), criteria


def _last_part(no):
    return int(no.rsplit("-", 1)[1])


def _check_sequence(label, nos):
    """末尾の番号が 1 から連続していること（重複・欠番がないこと）を確かめる。"""
    problems = []
    seen = sorted(_last_part(n) for n in nos)
    dups = sorted({n for n in seen if seen.count(n) > 1})
    if dups:
        problems.append(f"{label}: 番号の重複 {dups}")
    expected = list(range(1, max(seen) + 1)) if seen else []
    missing = sorted(set(expected) - set(seen))
    if missing:
        problems.append(f"{label}: 欠番 {missing}")
    return problems


def validate(requirements, criteria):
    """重複・欠番・ひも付けの不備を一覧で返す（空なら問題なし）。"""
    problems = []
    req_nos = [q["no"] for q in requirements]
    dup_reqs = sorted({n for n in req_nos if req_nos.count(n) > 1})
    if dup_reqs:
        problems.append(f"要求事項No.の重複: {dup_reqs}")
    by_category = {}
    for n in req_nos:
        by_category.setdefault(n.rsplit("-", 1)[0], []).append(n)
    for cat, nos in by_category.items():
        problems += _check_sequence(f"中分類{cat}の要求事項", nos)

    known = set(req_nos)
    crit_nos = [c["no"] for c in criteria]
    dup_crits = sorted({n for n in crit_nos if crit_nos.count(n) > 1})
    if dup_crits:
        problems.append(f"評価基準No.の重複: {dup_crits}")
    by_req = {}
    for c in criteria:
        if c["req_no"] not in known:
            problems.append(f"評価基準{c['no']}: 要求事項{c['req_no']}が存在しません")
        if c["no"].rsplit("-", 1)[0] != c["req_no"]:
            problems.append(f"評価基準{c['no']}: 番号が要求事項{c['req_no']}と対応しません")
        by_req.setdefault(c["req_no"], []).append(c["no"])
    for q in requirements:
        if q["no"] not in by_req:
            problems.append(f"要求事項{q['no']}: 評価基準がひとつもありません")
        elif q["star3"] != any(c["level"] == "★3" for c in criteria if c["req_no"] == q["no"]):
            problems.append(f"要求事項{q['no']}: ★3の印と★3評価基準の有無が一致しません")
    for req_no, nos in by_req.items():
        problems += _check_sequence(f"要求事項{req_no}の評価基準", nos)
    return problems


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def build(path, source_url, retrieved, also_published_at=None):
    cells, sheet_name = read_sheet(path)
    domains, requirements, criteria = convert(cells)
    problems = validate(requirements, criteria)
    if problems:
        raise ConversionError("取り込んだ内容に不備があります:\n  " + "\n  ".join(problems))
    star3_reqs = [q for q in requirements if q["star3"]]
    star3_crits = [c for c in criteria if c["level"] == "★3"]
    return {
        "source": {
            "title": "★3・★4要求事項及び評価基準",
            "publisher": "経済産業省",
            "url": source_url,
            "also_published_at": also_published_at or [],
            "retrieved": retrieved,
            "file_name": path.replace("\\", "/").split("/")[-1],
            "sha256": sha256_of(path),
            "sheet_name": sheet_name,
            "license": "公共データ利用規約（第1.0版）（PDL1.0）",
        },
        "counts": {
            "requirements_all": len(requirements),
            "criteria_all": len(criteria),
            "requirements_star3": len(star3_reqs),
            "criteria_star3": len(star3_crits),
        },
        "domains": domains,
        "requirements": requirements,
        "criteria": criteria,
    }


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("xlsx")
    p.add_argument("--source-url", required=True)
    p.add_argument("--retrieved", required=True, help="取得日 (YYYY-MM-DD)")
    p.add_argument("--also-published-at", action="append", default=[],
                   help="同一ファイル（ハッシュ一致）の別の公開先URL")
    p.add_argument("--out", required=True)
    a = p.parse_args(argv)
    data = build(a.xlsx, a.source_url, a.retrieved, a.also_published_at)
    with open(a.out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.write("\n")
    c = data["counts"]
    print(f"要求事項 全{c['requirements_all']}件（★3: {c['requirements_star3']}件） / "
          f"評価基準 全{c['criteria_all']}項目（★3: {c['criteria_star3']}項目）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
