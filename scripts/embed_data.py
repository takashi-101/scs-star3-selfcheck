"""data/requirements.json と samples/*.json を index.html に埋め込む。

index.html 内の次の2つの <script type="application/json"> の中身を置き換える。
    <script id="scs-requirements" ...>  … 要求事項・評価基準
    <script id="scs-samples" ...>       … 見本データ（ファイル名順）

使い方:
    python scripts/embed_data.py
"""
import glob
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def to_script_json(obj):
    # </script> で中身が途切れないよう "</" をエスケープする
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


def replace_block(html, block_id, payload):
    pattern = re.compile(
        r'(<script id="%s" type="application/json">)(.*?)(</script>)' % re.escape(block_id), re.S)
    if not pattern.search(html):
        raise SystemExit(f"index.html に <script id=\"{block_id}\"> がありません")
    return pattern.sub(lambda m: m.group(1) + payload + m.group(3), html, count=1)


def load_samples():
    samples = []
    for path in sorted(glob.glob(os.path.join(ROOT, "samples", "*.json"))):
        with open(path, encoding="utf-8") as f:
            samples.append(json.load(f))
    return samples


def main():
    html_path = os.path.join(ROOT, "index.html")
    with open(os.path.join(ROOT, "data", "requirements.json"), encoding="utf-8") as f:
        requirements = json.load(f)
    with open(html_path, encoding="utf-8") as f:
        html = f.read()
    html = replace_block(html, "scs-requirements", to_script_json(requirements))
    html = replace_block(html, "scs-samples", to_script_json(load_samples()))
    with open(html_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)
    print("index.html に埋め込みました")
    return 0


if __name__ == "__main__":
    sys.exit(main())
