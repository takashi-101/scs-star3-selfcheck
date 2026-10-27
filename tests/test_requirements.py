"""変換スクリプトと、取り込んだデータ・埋め込み・見本データの整合を確かめる。

実行: python -m unittest discover -s tests
"""
import io
import json
import os
import re
import sys
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import convert_requirements as conv  # noqa: E402
import embed_data  # noqa: E402

XLSX = os.path.join(ROOT, "data", "source", "scs_3.xlsx")
JSON_PATH = os.path.join(ROOT, "data", "requirements.json")
HTML_PATH = os.path.join(ROOT, "index.html")


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def embedded(block_id):
    with open(HTML_PATH, encoding="utf-8") as f:
        html = f.read()
    m = re.search(r'<script id="%s" type="application/json">(.*?)</script>' % block_id, html, re.S)
    return json.loads(m.group(1))


def make_xlsx(rows, merges=()):
    """テスト用の最小の xlsx をメモリ上に作る。rows は [[セルの値, ...], ...]（B列から）。"""
    def col(i):
        s = ""
        i += 2  # B列から
        while i:
            i, r = divmod(i - 1, 26)
            s = chr(65 + r) + s
        return s

    def esc(v):
        return v.replace("&", "&amp;").replace("<", "&lt;")

    data = []
    for r, row in enumerate(rows, start=2):
        cells = "".join(
            f'<c r="{col(i)}{r}" t="inlineStr"><is><t>{esc(v)}</t></is></c>'
            for i, v in enumerate(row) if v != "")
        data.append(f'<row r="{r}">{cells}</row>')
    merge_xml = ""
    if merges:
        merge_xml = '<mergeCells>%s</mergeCells>' % "".join(f'<mergeCell ref="{m}"/>' for m in merges)
    sheet = ('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<sheetData>{"".join(data)}</sheetData>{merge_xml}</worksheet>')
    wb = ('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="test" sheetId="1" r:id="rId1"/></sheets></workbook>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    buf.seek(0)
    return buf


HEADER = ["大分類No.", "大分類", "中分類No.", "中分類", "要求事項No.", "★3", "★4", "要求事項名",
          "要求事項", "★3/★4", "評価基準No.", "評価基準", "NIST CSFにおける機能"]


def row(req_no, crit_no, level="★3", title="T", text="X"):
    d, c = req_no.split("-")[0], "-".join(req_no.split("-")[:2])
    return [d, "分野", c, "中分類", req_no, "○" if level == "★3" else "", "○", title, text,
            level, crit_no, "基準" + crit_no, "識別(ID)"]


class OfficialDataTest(unittest.TestCase):
    """公式Excelから作った data/requirements.json についての確認。"""

    @classmethod
    def setUpClass(cls):
        cls.data = load_json(JSON_PATH)

    def test_regenerated_matches_committed(self):
        s = self.data["source"]
        fresh = conv.build(XLSX, s["url"], s["retrieved"], s["also_published_at"])
        self.assertEqual(fresh, self.data)

    def test_source_hash_matches_file(self):
        self.assertEqual(conv.sha256_of(XLSX), self.data["source"]["sha256"])
        self.assertEqual(len(self.data["source"]["sha256"]), 64)
        self.assertTrue(self.data["source"]["url"].startswith("https://www.meti.go.jp/"))
        self.assertRegex(self.data["source"]["retrieved"], r"^\d{4}-\d{2}-\d{2}$")

    def test_no_duplicate_or_missing_requirement_numbers(self):
        nos = [q["no"] for q in self.data["requirements"]]
        self.assertEqual(len(nos), len(set(nos)), "要求事項No.が重複している")
        by_cat = {}
        for n in nos:
            by_cat.setdefault(n.rsplit("-", 1)[0], []).append(int(n.rsplit("-", 1)[1]))
        for cat, seq in by_cat.items():
            self.assertEqual(sorted(seq), list(range(1, len(seq) + 1)), f"中分類{cat}に欠番がある")

    def test_no_duplicate_or_missing_criterion_numbers(self):
        nos = [c["no"] for c in self.data["criteria"]]
        self.assertEqual(len(nos), len(set(nos)), "評価基準No.が重複している")
        by_req = {}
        for n in nos:
            by_req.setdefault(n.rsplit("-", 1)[0], []).append(int(n.rsplit("-", 1)[1]))
        for req, seq in by_req.items():
            self.assertEqual(sorted(seq), list(range(1, len(seq) + 1)), f"要求事項{req}の評価基準に欠番がある")

    def test_every_criterion_belongs_to_a_requirement(self):
        reqs = {q["no"] for q in self.data["requirements"]}
        for c in self.data["criteria"]:
            self.assertIn(c["req_no"], reqs, c["no"])
            self.assertEqual(c["no"].rsplit("-", 1)[0], c["req_no"], c["no"])
            self.assertTrue(c["text"].strip(), c["no"])

    def test_every_star3_requirement_has_star3_criteria(self):
        for q in self.data["requirements"]:
            has3 = any(c["level"] == "★3" for c in self.data["criteria"] if c["req_no"] == q["no"])
            self.assertEqual(q["star3"], has3, q["no"])

    def test_counts_are_consistent(self):
        c = self.data["counts"]
        self.assertEqual(c["requirements_star3"], sum(q["star3"] for q in self.data["requirements"]))
        self.assertEqual(c["criteria_star3"], sum(x["level"] == "★3" for x in self.data["criteria"]))
        self.assertEqual(conv.validate(self.data["requirements"], self.data["criteria"]), [])

    def test_no_phonetic_text_left(self):
        # 元のExcelには ふりがな（<rPh>）が付いたセルがある。それが本文に混ざっていないこと。
        with zipfile.ZipFile(XLSX) as z:
            raw = z.read("xl/sharedStrings.xml").decode("utf-8")
        phonetics = set(re.findall(r"<rPh[^>]*><t>([^<]+)</t></rPh>", raw))
        self.assertTrue(phonetics, "ふりがなを含むセルが見つからない（前提が変わった）")
        texts = [c["text"] for c in self.data["criteria"]] + [q["text"] for q in self.data["requirements"]]
        for t in texts:
            for p in phonetics:
                if len(p) >= 4 and re.fullmatch(r"[ァ-ヶー・]+", p):
                    self.assertFalse(t.endswith(p), f"ふりがな {p!r} が混入: {t[:30]}")

    def test_phonetic_runs_are_ignored(self):
        si = conv.ET.fromstring(
            '<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<r><t>社内</t></r><r><t>ルール</t></r><rPh sb="0" eb="2"><t>シャナイ</t></rPh></si>')
        self.assertEqual(conv._text_of(si), "社内ルール")


class ConverterUnitTest(unittest.TestCase):
    """小さな表を作って、変換と検査の動きを確かめる。"""

    def convert(self, rows, merges=()):
        cells, _ = conv.read_sheet(make_xlsx([HEADER] + rows, merges))
        return conv.convert(cells)

    def test_basic(self):
        _, reqs, crits = self.convert([row("1-1-1", "1-1-1-1"), row("1-1-1", "1-1-1-2", "★4")])
        self.assertEqual([q["no"] for q in reqs], ["1-1-1"])
        self.assertTrue(reqs[0]["star3"])
        self.assertEqual([c["level"] for c in crits], ["★3", "★4"])
        self.assertEqual(conv.validate(reqs, crits), [])

    def test_star3_flag_taken_from_any_row(self):
        # 先頭行が★4だけの評価基準でも、要求事項は★3の対象になる
        _, reqs, _ = self.convert([row("1-1-1", "1-1-1-1", "★4"), row("1-1-1", "1-1-1-2")])
        self.assertTrue(reqs[0]["star3"])

    def test_merged_cells_are_filled(self):
        r1 = row("1-1-1", "1-1-1-1")
        r2 = row("1-1-1", "1-1-1-2")
        for i in range(0, 10):  # B〜K列を空にして結合で埋める
            r2[i] = ""
        merges = [f"{chr(66 + i)}3:{chr(66 + i)}4" for i in range(10)]
        _, reqs, crits = self.convert([r1, r2], merges)
        self.assertEqual(len(reqs), 1)
        self.assertEqual([c["req_no"] for c in crits], ["1-1-1", "1-1-1"])

    def test_detects_duplicate_criterion(self):
        _, reqs, crits = self.convert([row("1-1-1", "1-1-1-1"), row("1-1-1", "1-1-1-1")])
        self.assertTrue(any("重複" in p for p in conv.validate(reqs, crits)))

    def test_detects_missing_criterion_number(self):
        _, reqs, crits = self.convert([row("1-1-1", "1-1-1-1"), row("1-1-1", "1-1-1-3")])
        self.assertTrue(any("欠番" in p for p in conv.validate(reqs, crits)))

    def test_detects_missing_requirement_number(self):
        _, reqs, crits = self.convert([row("1-1-1", "1-1-1-1"), row("1-1-3", "1-1-3-1")])
        self.assertTrue(any("欠番" in p for p in conv.validate(reqs, crits)))

    def test_detects_criterion_not_matching_requirement(self):
        _, reqs, crits = self.convert([row("1-1-1", "1-1-2-1")])
        problems = conv.validate(reqs, crits)
        self.assertTrue(any("対応しません" in p for p in problems))

    def test_detects_orphan_criterion(self):
        reqs = [{"no": "1-1-1", "star3": True}]
        crits = [{"no": "1-1-1-1", "req_no": "1-1-1", "level": "★3"},
                 {"no": "9-9-9-1", "req_no": "9-9-9", "level": "★3"}]
        self.assertTrue(any("存在しません" in p for p in conv.validate(reqs, crits)))

    def test_rejects_conflicting_requirement_text(self):
        with self.assertRaises(conv.ConversionError):
            self.convert([row("1-1-1", "1-1-1-1", text="A"), row("1-1-1", "1-1-1-2", text="B")])

    def test_rejects_missing_header(self):
        with self.assertRaises(conv.ConversionError):
            conv.convert(conv.read_sheet(make_xlsx([["a", "b"]]))[0])


class EmbeddedAndSamplesTest(unittest.TestCase):
    """index.html への埋め込みと見本データの確認。"""

    def test_html_embeds_current_requirements(self):
        self.assertEqual(embedded("scs-requirements"), load_json(JSON_PATH))

    def test_html_embeds_current_samples(self):
        self.assertEqual(embedded("scs-samples"), embed_data.load_samples())

    def test_html_loads_nothing_from_outside(self):
        with open(HTML_PATH, encoding="utf-8") as f:
            html = f.read()
        self.assertIn("connect-src 'none'", html)
        self.assertNotRegex(html, r'<script[^>]+src=', "外部スクリプトを読み込んでいる")
        self.assertNotRegex(html, r'<link[^>]+href="(?!data:)', "外部の資源を読み込んでいる")
        self.assertNotRegex(html, r'(src|href)="(https?:)?//', "外部の資源を参照している")
        self.assertNotRegex(html, r'\b(fetch|XMLHttpRequest|WebSocket|EventSource|sendBeacon)\s*\(', "通信する処理がある")
        self.assertNotRegex(html, r'url\(\s*["\']?https?:', "CSSで外部の資源を読み込んでいる")

    def test_samples_are_valid(self):
        star3 = {c["no"] for c in load_json(JSON_PATH)["criteria"] if c["level"] == "★3"}
        sha = load_json(JSON_PATH)["source"]["sha256"]
        samples = embed_data.load_samples()
        self.assertEqual(len(samples), 2)
        kinds = sorted(s["org"]["kind"] for s in samples)
        self.assertEqual(kinds, ["中小企業", "地方公共団体"])
        for s in samples:
            name = s["org"]["name"]
            self.assertIn("架空", name)
            self.assertEqual(s["format"], "scs-star3-selfcheck")
            self.assertEqual(s["requirementsSha256"], sha)
            self.assertLessEqual(set(s["checks"]), star3, name)
            for v in s["checks"].values():
                self.assertIn(v["status"], ["done", "partial", "todo", "na", ""])
                if v["status"] == "na":
                    self.assertTrue(v["memo"].strip(), f"{name}: 対象外の理由が空")
            ids = [x["id"] for x in s["services"]]
            self.assertEqual(len(ids), len(set(ids)))
            for f in s["fallbacks"]:
                self.assertTrue(f["serviceIds"], f["operation"])
                for sid in f["serviceIds"]:
                    self.assertIn(sid, ids, f"{name}: 切替表が台帳にないサービスを指している")


if __name__ == "__main__":
    unittest.main()
